<#
.SYNOPSIS
    Enquête sur la "disparition" de mails dans une boîte Exchange Online :
    qui a supprimé ou déplacé quoi, quand, et depuis quel client.

.DESCRIPTION
    Enchaîne les vérifications utiles pour ce type de demande :
      1. Contexte de la boîte : type, audit, règles de boîte de réception,
         règles de rangement (sweep), comptes ayant des droits dessus.
      2. Trace de transport du ou des mails recherchés (si -Subject est fourni).
      3. Journal d'audit unifié : suppressions et déplacements sur la boîte,
         récupérés par pagination (pas de troncature à 5000 résultats).
      4. Synthèse : événements concernant le mail recherché, puis répartition
         par compte / opération / client sur toute la période.

    Le script est en lecture seule : il ne modifie rien dans le tenant.

    Limites de rétention à connaître :
      - trace de messages : 90 jours ;
      - journal d'audit : selon la licence (180 jours en audit standard).

.PARAMETER Mailbox
    Boîte à examiner (adresse, alias ou UPN).

.PARAMETER StartDate
    Début de la période, en heure locale. Par défaut : il y a 7 jours.
    Utiliser le format ISO ("2026-10-05" ou "2026-10-05 09:00") : PowerShell
    lit "05/10/2026" à l'américaine (10 mai).

.PARAMETER EndDate
    Fin de la période, en heure locale. Par défaut : maintenant.

.PARAMETER Subject
    Fragment du sujet du mail recherché (sans jokers). Facultatif : sans lui,
    le script donne uniquement la vue d'ensemble de la boîte.

.PARAMETER Operations
    Opérations d'audit recherchées.
    Par défaut : MoveToDeletedItems, SoftDelete, HardDelete, Move.

.PARAMETER SliceHours
    Taille des tranches de recherche dans l'audit, en heures (défaut : 24).
    À réduire si le script signale qu'une tranche atteint 50 000 événements.

.PARAMETER ExportCsv
    Chemin d'un fichier CSV où exporter tous les événements de la boîte.

.PARAMETER SkipTrace
    Ne pas interroger la trace de messages.

.PARAMETER PassThru
    Renvoie les événements dans le pipeline (pour un traitement ultérieur).

.EXAMPLE
    .\Get-MailboxDeletionAudit.ps1 -Mailbox barbes@crous-paris.fr `
        -StartDate "2026-10-05" -EndDate "2026-10-09" -Subject "désinsectisation"

    Retrace le parcours d'un mail précis et donne la vue d'ensemble de la boîte.

.EXAMPLE
    .\Get-MailboxDeletionAudit.ps1 -Mailbox barbes@crous-paris.fr -ExportCsv .\barbes.csv

    Vue d'ensemble des 7 derniers jours, avec export CSV du détail.

.EXAMPLE
    $ev = .\Get-MailboxDeletionAudit.ps1 -Mailbox barbes@crous-paris.fr -PassThru
    $ev | Where-Object Utilisateur -like "prenom.nom*" | Format-Table Date, Operation, Sujet

.NOTES
    Prérequis : module ExchangeOnlineManagement et un rôle donnant accès à
    Search-UnifiedAuditLog (Audit Logs ou View-Only Audit Logs).
    Lecture du champ "Acces" : Propriétaire = le titulaire de la boîte,
    Délégué = un compte ayant des droits dessus, Admin = accès de type
    administrateur.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Mailbox,

    [datetime]$StartDate = (Get-Date).Date.AddDays(-7),

    [datetime]$EndDate = (Get-Date),

    [string]$Subject,

    [string[]]$Operations = @('MoveToDeletedItems', 'SoftDelete', 'HardDelete', 'Move'),

    [ValidateRange(1, 24)]
    [int]$SliceHours = 24,

    [string]$ExportCsv,

    [switch]$SkipTrace,

    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

function Write-Section {
    param([string]$Title)
    Write-Host ''
    Write-Host ('=== {0} ===' -f $Title) -ForegroundColor Cyan
}

if ($EndDate -le $StartDate) {
    throw "EndDate ($EndDate) doit être postérieure à StartDate ($StartDate)."
}

# Exchange Online raisonne en UTC : on convertit une fois pour toutes.
$startUtc = $StartDate.ToUniversalTime()
$endUtc   = $EndDate.ToUniversalTime()

# --- Connexion ---------------------------------------------------------------
$connected = $false
if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
    $connected = [bool](Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })
}
if (-not $connected) {
    Write-Host 'Connexion à Exchange Online...'
    Connect-ExchangeOnline -ShowBanner:$false
}

# --- 1. Contexte de la boîte -------------------------------------------------
Write-Section 'Boîte'
$mbx  = Get-Mailbox -Identity $Mailbox
$upn  = $mbx.UserPrincipalName
$smtp = [string]$mbx.PrimarySmtpAddress

$mbx | Format-List DisplayName, PrimarySmtpAddress, UserPrincipalName, RecipientTypeDetails, AuditEnabled | Out-Host
Write-Host ('Période : du {0:dd/MM/yyyy HH:mm} au {1:dd/MM/yyyy HH:mm} (heure locale)' -f $StartDate, $EndDate)

Write-Section 'Règles de boîte de réception'
$rules = @(Get-InboxRule -Mailbox $upn -IncludeHidden)
if ($rules.Count) {
    $rules | Format-List Name, Enabled, Priority, DeleteMessage, MoveToFolder, Description | Out-Host
} else {
    Write-Host 'Aucune.'
}

Write-Section 'Règles de rangement (sweep)'
$sweep = @(Get-SweepRule -Mailbox $upn)
if ($sweep.Count) {
    $sweep | Format-List Name, Enabled, Sender, KeepForDays, KeepLatest, DestinationFolder | Out-Host
} else {
    Write-Host 'Aucune.'
}

Write-Section 'Comptes ayant des droits sur la boîte'
$perms = @(Get-MailboxPermission -Identity $upn |
    Where-Object { -not $_.IsInherited -and $_.User -ne 'NT AUTHORITY\SELF' })
if ($perms.Count) {
    $perms | Format-Table User, AccessRights, Deny -AutoSize | Out-Host
} else {
    Write-Host 'Aucun (hors titulaire).'
}

# --- 2. Trace de messages ----------------------------------------------------
$traceIds = @()
if ($Subject -and -not $SkipTrace) {
    Write-Section ('Trace de messages (sujet contenant "{0}")' -f $Subject)

    $traceStart = $startUtc
    $limit = (Get-Date).ToUniversalTime().AddDays(-90)
    if ($traceStart -lt $limit) {
        Write-Warning 'La trace de messages ne remonte pas au-delà de 90 jours : début de trace ajusté.'
        $traceStart = $limit
    }

    try {
        # Une requête de trace couvre 10 jours au plus : on découpe.
        $cursor = $traceStart
        $traces = @(while ($cursor -lt $endUtc) {
            $next = $cursor.AddDays(10)
            if ($next -gt $endUtc) { $next = $endUtc }
            Get-MessageTraceV2 -RecipientAddress $smtp -StartDate $cursor -EndDate $next -ResultSize 5000
            $cursor = $next
        })

        $traces = @($traces | Where-Object { $_.Subject -like "*$Subject*" } | Sort-Object Received)
        if ($traces.Count) {
            $traceIds = @($traces | ForEach-Object { $_.MessageId })
            $traces | Select-Object `
                @{ n = 'Reçu'; e = { [datetime]::SpecifyKind([datetime]$_.Received, [DateTimeKind]::Utc).ToLocalTime() } },
                SenderAddress, Subject, Status, MessageId |
                Format-List | Out-Host
        } else {
            Write-Host 'Aucun mail correspondant dans la trace sur cette période.'
        }
    } catch {
        Write-Warning ('Trace de messages indisponible : {0}' -f $_.Exception.Message)
    }
}

# --- 3. Journal d'audit unifié (paginé) --------------------------------------
Write-Section "Journal d'audit"
$raw    = New-Object 'System.Collections.Generic.List[object]'
$cursor = $startUtc
$slice  = 0

while ($cursor -lt $endUtc) {
    $next = $cursor.AddHours($SliceHours)
    if ($next -gt $endUtc) { $next = $endUtc }

    $sid   = 'mbxaudit-{0}-{1}' -f $slice, (Get-Random)
    $count = 0
    do {
        $lot = @(Search-UnifiedAuditLog -StartDate $cursor -EndDate $next `
                -RecordType ExchangeItemGroup -Operations $Operations `
                -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000)
        if ($lot.Count) {
            $raw.AddRange($lot)
            $count += $lot.Count
        }
    } while ($lot.Count -gt 0)

    Write-Host ('  {0:dd/MM HH:mm} -> {1:dd/MM HH:mm} : {2} événements (tenant entier)' -f `
            $cursor.ToLocalTime(), $next.ToLocalTime(), $count)
    if ($count -ge 50000) {
        Write-Warning 'Tranche saturée (50 000) : résultats incomplets, relancer avec un -SliceHours plus petit.'
    }

    $cursor = $next
    $slice++
}

# --- 4. Filtrage sur la boîte, dédoublonnage, mise à plat ---------------------
$logonTypes = @{ '0' = 'Propriétaire'; '1' = 'Admin'; '2' = 'Délégué' }
$seen = @{}

$events = @($(foreach ($rec in $raw) {
    # Pré-filtre rapide avant de décoder le JSON
    if ($rec.AuditData -notlike "*$upn*") { continue }

    $d = $rec.AuditData | ConvertFrom-Json
    if ($d.MailboxOwnerUPN -ne $upn) { continue }
    if ($seen.ContainsKey([string]$d.Id)) { continue }   # ReturnLargeSet renvoie des doublons
    $seen[[string]$d.Id] = $true

    $utc   = [datetime]::SpecifyKind([datetime]$d.CreationTime, [DateTimeKind]::Utc)
    $acces = $logonTypes[[string]$d.LogonType]
    if (-not $acces) { $acces = [string]$d.LogonType }

    $items = @($d.AffectedItems)
    if (-not $items.Count) { $items = @($null) }

    foreach ($item in $items) {
        $depuis = $d.Folder.Path
        if (-not $depuis -and $item) { $depuis = $item.ParentFolder.Path }

        $msgId = $null
        $sujet = $null
        if ($item) {
            $msgId = $item.InternetMessageId
            $sujet = $item.Subject
        }

        [pscustomobject]@{
            Date              = $utc.ToLocalTime()
            Operation         = $d.Operation
            Utilisateur       = $d.UserId
            Acces             = $acces
            Client            = $d.ClientInfoString
            IP                = $d.ClientIPAddress
            Depuis            = $depuis
            Vers              = $d.DestFolder.Path
            Sujet             = $sujet
            MailDeLaTrace     = [bool]($msgId -and ($traceIds -contains $msgId))
            InternetMessageId = $msgId
            DateUTC           = $utc
            AuditId           = $d.Id
        }
    }
}) | Sort-Object Date)

Write-Host ''
Write-Host ('{0} élément(s) supprimé(s) ou déplacé(s) dans la boîte {1} sur la période.' -f $events.Count, $smtp)

# --- 5. Restitution ----------------------------------------------------------
if ($Subject) {
    Write-Section ('Événements sur les mails dont le sujet contient "{0}"' -f $Subject)
    $hits = @($events | Where-Object { $_.Sujet -like "*$Subject*" })
    if ($hits.Count) {
        $hits | Format-List Date, Operation, Utilisateur, Acces, Client, IP, Depuis, Vers, Sujet, MailDeLaTrace | Out-Host
    } else {
        Write-Host 'Aucun événement audité pour ce sujet sur la période.'
        Write-Host "Pistes : élargir la période, vérifier le sujet, ou le mail n'a pas été supprimé/déplacé."
    }
}

Write-Section 'Vue d''ensemble : qui fait quoi sur cette boîte'
if ($events.Count) {
    $events | Group-Object Utilisateur, Acces, Operation, Client |
        Sort-Object Count -Descending |
        Select-Object @{ n = 'Eléments'; e = { $_.Count } },
                      @{ n = 'Utilisateur'; e = { $_.Group[0].Utilisateur } },
                      @{ n = 'Acces'; e = { $_.Group[0].Acces } },
                      @{ n = 'Operation'; e = { $_.Group[0].Operation } },
                      @{ n = 'Client'; e = { $_.Group[0].Client } } |
        Format-Table -AutoSize -Wrap | Out-Host
} else {
    Write-Host 'Rien à afficher.'
}

if ($ExportCsv) {
    $events | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host ('Détail exporté dans {0}' -f $ExportCsv)
}

if ($PassThru) {
    $events
}