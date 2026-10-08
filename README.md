# Get-MailboxDeletionAudit

Script PowerShell d'investigation pour Exchange Online : il répond à la question « qui a supprimé ou déplacé ce mail, quand, et depuis quel appareil ? ».

Cas d'usage typique : un utilisateur signale que des mails « disparaissent tout seuls » de sa boîte de réception. Le script vérifie les règles de la boîte, retrouve le mail dans la trace de transport, puis interroge le journal d'audit pour identifier le compte et le client à l'origine du déplacement.

Le script est en **lecture seule** : il ne modifie rien dans le tenant.

## Ce que fait le script

1. **Contexte de la boîte** : type, état de l'audit, règles de boîte de réception (y compris masquées), règles de rangement (sweep), comptes ayant des droits dessus.
2. **Trace de messages** : retrouve le ou les mails dont le sujet contient le texte recherché, avec l'heure de remise.
3. **Journal d'audit unifié** : récupère les suppressions et déplacements de la période, par tranches paginées, puis ne garde que ceux de la boîte examinée.
4. **Restitution** : les événements portant sur le mail recherché, puis une vue d'ensemble de qui fait quoi sur la boîte.

## Prérequis

- Windows PowerShell 5.1 ou PowerShell 7.
- Module `ExchangeOnlineManagement` en version 3.7.0 ou plus (nécessaire pour `Get-MessageTraceV2`) :

  ```powershell
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  ```

- Un compte disposant :
  - d'un accès au journal d'audit (rôle *Audit Logs* ou *View-Only Audit Logs*) ;
  - des droits de lecture sur les boîtes, leurs règles et leurs permissions ;
  - de l'accès à la trace de messages.
- L'audit doit être activé sur le tenant et sur la boîte (c'est le cas par défaut).

Si aucune session Exchange Online n'est ouverte, le script lance `Connect-ExchangeOnline`.

## Utilisation

Retracer le parcours d'un mail précis :

```powershell
.\Get-MailboxDeletionAudit.ps1 -Mailbox accueil@exemple.fr `
    -StartDate "2026-10-05" -EndDate "2026-10-09" -Subject "facture"
```

Vue d'ensemble des 7 derniers jours, avec export du détail :

```powershell
.\Get-MailboxDeletionAudit.ps1 -Mailbox accueil@exemple.fr -ExportCsv .\accueil.csv
```

Récupérer les événements pour les retraiter :

```powershell
$ev = .\Get-MailboxDeletionAudit.ps1 -Mailbox accueil@exemple.fr -PassThru
$ev | Where-Object Utilisateur -like "jean.dupont*" |
    Format-Table Date, Operation, Sujet
```

L'aide intégrée est disponible avec `Get-Help .\Get-MailboxDeletionAudit.ps1 -Full`.

## Paramètres

| Paramètre | Obligatoire | Défaut | Rôle |
|---|---|---|---|
| `-Mailbox` | oui | | Boîte à examiner (adresse, alias ou UPN). |
| `-StartDate` | non | il y a 7 jours | Début de la période, en heure locale. |
| `-EndDate` | non | maintenant | Fin de la période, en heure locale. |
| `-Subject` | non | | Fragment du sujet recherché, sans jokers. Sans lui, seule la vue d'ensemble est produite. |
| `-Operations` | non | `MoveToDeletedItems`, `SoftDelete`, `HardDelete`, `Move` | Opérations d'audit recherchées. |
| `-SliceHours` | non | `24` | Taille des tranches de recherche dans l'audit, de 1 à 24 heures. |
| `-ExportCsv` | non | | Fichier CSV où exporter tous les événements de la boîte (séparateur `;`). |
| `-SkipTrace` | non | | Ne pas interroger la trace de messages. |
| `-PassThru` | non | | Renvoie les événements dans le pipeline. |

> **Format des dates** : utiliser le format ISO (`"2026-10-05"` ou `"2026-10-05 09:00"`). PowerShell interprète `"05/10/2026"` à l'américaine, soit le 10 mai.

## Exemple de résultat

Les données ci-dessous sont fictives.

```text
=== Boîte ===

DisplayName          : Accueil
PrimarySmtpAddress   : accueil@exemple.fr
UserPrincipalName    : accueil@exemple.fr
RecipientTypeDetails : UserMailbox
AuditEnabled         : True

Période : du 05/10/2026 00:00 au 09/10/2026 00:00 (heure locale)

=== Règles de boîte de réception ===

Name          : Junk E-mail Rule
Enabled       : True
Priority      : 0
DeleteMessage : False
MoveToFolder  :
Description   :

=== Règles de rangement (sweep) ===
Aucune.

=== Comptes ayant des droits sur la boîte ===

User                    AccessRights Deny
----                    ------------ ----
jean.dupont@exemple.fr  {FullAccess} False
marie.martin@exemple.fr {FullAccess} False

=== Trace de messages (sujet contenant "facture") ===

Reçu          : 05/10/2026 10:01:20
SenderAddress : fournisseur@exemple.com
Subject       : Facture octobre
Status        : Delivered
MessageId     : <abc123@mail.exemple.com>

=== Journal d'audit ===
  05/10 00:00 -> 06/10 00:00 : 6480 événements (tenant entier)
  ...

42 élément(s) supprimé(s) ou déplacé(s) dans la boîte accueil@exemple.fr sur la période.

=== Événements sur les mails dont le sujet contient "facture" ===

Date          : 05/10/2026 10:59:30
Operation     : MoveToDeletedItems
Utilisateur   : jean.dupont@exemple.fr
Acces         : Délégué
Client        : Client=OutlookService;Outlook-Android/2.0;
IP            : 203.0.113.25
Depuis        : \Inbox
Vers          : \Deleted Items
Sujet         : Facture octobre
MailDeLaTrace : True

=== Vue d'ensemble : qui fait quoi sur cette boîte ===

Eléments Utilisateur             Acces        Operation          Client
-------- -----------             -----        ---------          ------
      31 jean.dupont@exemple.fr  Délégué      MoveToDeletedItems Client=OutlookService;Outlook-Android/2.0;
       8 accueil@exemple.fr      Propriétaire SoftDelete         Client=OWA;Action=ViaProxy
       3 marie.martin@exemple.fr Délégué      MoveToDeletedItems Client=MSExchangeRPC
```

## Lire le résultat

**Colonne `Acces`**

| Valeur | Signification |
|---|---|
| Propriétaire | Action faite en étant connecté avec le compte de la boîte. |
| Délégué | Action faite par un autre compte ayant des droits sur la boîte. |
| Admin | Accès de type administrateur. |

**Colonne `Operation`**

| Valeur | Signification |
|---|---|
| `MoveToDeletedItems` | Mail envoyé dans la corbeille. |
| `SoftDelete` | Mail supprimé définitivement côté utilisateur (vidage de corbeille, Maj+Suppr), encore récupérable. |
| `HardDelete` | Mail purgé des éléments récupérables. |
| `Move` | Déplacement vers un autre dossier (voir les limites). |

**Colonne `Client`** : elle identifie le logiciel utilisé. Quelques valeurs courantes : `Client=OWA` (Outlook sur le web), `Client=MSExchangeRPC` (Outlook pour Windows), `Client=OutlookService;Outlook-Android` ou `Outlook-iOS` (application mobile).

**Colonne `MailDeLaTrace`** : `True` quand l'événement porte exactement sur un mail trouvé dans la trace (même identifiant de message). C'est ce qui permet d'affirmer qu'il s'agit bien du mail signalé, et pas d'un autre de même sujet.

**Interprétation**

- Un déplacement dans la seconde qui suit la remise évoque un automatisme ; un déplacement plusieurs minutes ou heures après, depuis un client identifié, évoque une action d'un utilisateur ou d'un de ses appareils.
- Le journal donne le compte et l'appareil, pas l'intention. Sur mobile, un simple balayage suffit à envoyer un message, ou une conversation entière, à la corbeille.
- Sur une boîte utilisée à plusieurs, la suppression faite par l'un s'applique à tous.

## Limites

- **Rétention** : la trace de messages remonte à 90 jours ; le journal d'audit dépend de la licence (180 jours en audit standard).
- **Règles serveur** : un mail déplacé par une règle de boîte de réception au moment de la remise n'apparaît pas comme une action d'utilisateur dans l'audit. C'est pourquoi le script affiche les règles : une règle de suppression ou de déplacement explique à elle seule ce cas.
- **Opération `Move`** : elle n'est pas auditée par défaut. Les déplacements vers un dossier autre que la corbeille n'apparaissent que si elle a été ajoutée aux actions auditées de la boîte.
- **Volume** : la recherche d'audit porte sur tout le tenant avant filtrage, par tranches limitées à 50 000 événements. Si le script signale une tranche saturée, relancer avec un `-SliceHours` plus petit (par exemple `6`).
- **Durée** : quelques minutes sur un tenant de taille moyenne, davantage sur une longue période.
- **Délai de l'audit** : un événement peut mettre de 30 minutes à plusieurs heures à apparaître dans le journal.

## Notes techniques

- Les dates sont saisies et affichées en heure locale ; le script convertit en UTC pour interroger Exchange Online.
- La recherche d'audit utilise `Search-UnifiedAuditLog` avec `-SessionCommand ReturnLargeSet`. Ce mode renvoie des doublons, que le script élimine à partir de l'identifiant de l'enregistrement.
- `Search-MailboxAuditLog` n'est pas utilisée : Microsoft a retiré cette cmdlet.
- Le fichier est encodé en UTF-8 avec BOM. Conserver cet encodage en cas de modification, sans quoi Windows PowerShell 5.1 lit mal les caractères accentués.

## Bon usage

Le résultat désigne nommément des personnes. Il est à réserver aux demandes légitimes d'investigation et à communiquer dans le cadre prévu par les règles internes de l'organisation.
