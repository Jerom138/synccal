# Synchronisation de trois agendas

Deux agendas Microsoft 365 d'organisations distinctes et un agenda Google, maintenus
alignés sans doublon.

Script de synchronisation à trois agendas, relançable à volonté et sans doublon.
Chaque événement créé dans un agenda est recopié dans les deux autres.

```bash
python sync.py            # simulation : montre ce qui serait fait, n'écrit rien
python sync.py --apply    # applique
python sync.py --apply -v # applique, en détaillant chaque action
python sync.py --purge    # simule le retrait de toutes les copies
```

Le script est **idempotent** : le relancer sans rien changer ne produit aucune action.
C'est cette propriété qui garantit l'absence de doublon.

---

## Comment ça marche

Chaque copie porte un marqueur invisible (`SyncOrigin`) qui désigne son événement
d'origine, et une empreinte du contenu (`SyncHash`). À chaque exécution, le script
relit les trois agendas, reconstruit l'état depuis ces marqueurs et en déduit les
créations, modifications et suppressions nécessaires.

Il n'existe **aucune base de données locale** : toute la mémoire de la synchronisation
vit dans les agendas eux-mêmes. Le script survit donc à un changement de poste, à une
suppression manuelle ou à une exécution interrompue.

Deux règles coupent les boucles et les doublons :

- un événement porteur d'un marqueur n'est **jamais** repris comme source ;
- deux copies désignant la même origine ⇒ la surnuméraire est supprimée.

### Garde-fous

| Protection | Comportement |
|---|---|
| Lecture en échec | Le run s'arrête **sans rien supprimer**. Un agenda illisible n'est jamais confondu avec un agenda vide. |
| Suppressions massives | Au-delà de `safety.max_deletions` (20), arrêt et demande de `--force`. |
| Agenda momentanément vide | Un agenda qui ne renvoie plus aucun événement alors que des copies le désignent ailleurs arrête le run. C'est la signature d'un profil Outlook en cours de première synchronisation ou d'un OST reconstruit — appliquer le plan effacerait toutes ses copies. |
| Agenda momentanément vide | Un agenda qui ne renvoie plus aucun événement alors que des copies le désignent ailleurs arrête le run. C'est la signature d'un profil Outlook en cours de première synchronisation ou d'un OST reconstruit — appliquer le plan effacerait toutes ses copies. |
| Copies hors fenêtre | Jamais supprimées : leur original est peut-être simplement hors de portée. |
| Agenda source désactivé | Ses copies sont laissées intactes dans les autres agendas. |
| Invitations | Les copies n'ont **aucun participant** et ne sont pas des réunions : aucune invitation n'est jamais envoyée. |
| Rappels | Désactivés sur les copies — l'original porte déjà le sien. |
| Réunions annulées | Outlook garde dans l'agenda les réunions annulées (barrées, préfixées « Annulé: »). Elles ne sont pas recopiées. |
| Copie désynchronisée | Les horaires d'une copie sont comparés à ceux de sa source à chaque run, en plus de l'empreinte. Une copie retouchée à la main, ou mal écrite, est détectée et réalignée — l'empreinte seule ne suffirait pas, puisqu'elle décrit la source et resterait valide. |
| Copie désynchronisée | Les horaires d'une copie sont comparés à ceux de sa source à chaque run, en plus de l'empreinte. Une copie retouchée à la main, ou mal écrite, est détectée et réalignée — l'empreinte seule ne suffirait pas, puisqu'elle décrit la source et resterait valide. |

---

## Installation

```bash
pip install -r requirements.txt
copy config.example.yaml config.yaml
```

`config.yaml` n'est pas versionné : il contient vos adresses et les noms de vos calendriers.
Les jetons, journaux et sauvegardes d'agenda ne le sont pas davantage.

### 1. Google Agenda

1. Créer un projet sur [console.cloud.google.com](https://console.cloud.google.com) et
   activer **Google Calendar API**.
2. Créer un identifiant OAuth de type **Application de bureau** et déposer le JSON
   téléchargé dans `secrets/`. Son nom n'a pas d'importance : tout fichier
   `client_secret*.json` y est détecté automatiquement.
3. Créer le calendrier Google dédié dans Google Agenda.
4. Autoriser l'accès, une fois :

   ```bash
   python -u authorize_google.py
   ```

   Le script affiche l'URL d'autorisation, enregistre le jeton dans
   `secrets/token_google.json`, puis **liste vos calendriers Google** pour confirmer que
   le nom configuré correspond bien.

> **Statut de l'écran de consentement.** Deux régimes, à choisir en connaissance de cause :
>
> | Statut | Conséquence |
> |---|---|
> | **Test** | Seuls les comptes inscrits comme testeurs ont accès, et Google **invalide le jeton tous les 7 jours**. Il faut relancer `authorize_google.py` chaque semaine, sinon la synchro s'arrête. |
> | **En production** | Pas de liste de testeurs, jeton valide jusqu'à révocation. L'avertissement « application non vérifiée » subsiste, sans conséquence en usage personnel. |
>
> La configuration actuelle est en **Test**. `sync.py` ne lance jamais d'autorisation
> interactive : à l'expiration, il s'arrête avec un message explicite plutôt que de rester
> bloqué sur une invite — ce qui serait fatal pour une tâche planifiée.

### 2. Agendas Outlook

```bash
python check_outlook.py
```

Ce diagnostic liste les agendas joignables et affiche, pour chacun, la valeur exacte à
reporter dans le champ `store` de `config.yaml`.

- **Aucun profil Outlook classique ?** Lancer `OUTLOOK.EXE` une fois et créer un profil
  avec le compte Pro A. Le nouvel Outlook reste utilisable en parallèle.
- **Compte Pro B ajoutable ?** `Fichier ▸ Ajouter un compte`. Si l'ajout réussit, passer
  `pro_b.backend` à `outlook_com` dans `config.yaml` et renseigner son `store` : la
  synchronisation devient complète dans les deux sens.
- **Ajout refusé ?** Conserver `pro_b.backend: ics_readonly` et publier l'agenda depuis
  Outlook Web : `Paramètres ▸ Calendrier ▸ Calendriers partagés ▸ Publier un calendrier`,
  niveau **« Tous les détails »**. Coller le lien `.ics` dans `config.yaml`.

### 3. Vérifier

```bash
python sync.py
```

Lire le plan proposé. S'il est cohérent, relancer avec `--apply`.

> **Premier run : commencez petit.** Un agenda professionnel chargé contient vite
> plusieurs centaines d'événements sur un an, et chacun produit deux copies. Avant
> le run complet, réduisez temporairement la fenêtre pour valider le résultat sur un
> volume maîtrisable :
>
> ```yaml
> window:
>   past_days: 7
>   future_months: 1
> ```
>
> Vérifiez les copies dans les trois agendas, puis remettez `30` / `12`. Les runs
> suivants ne traitent que les écarts : seul le premier est long.

---

## Configuration

Tout se règle dans `config.yaml` :

| Clé | Rôle |
|---|---|
| `window.past_days` | profondeur dans le passé (30 par défaut) |
| `window.future_months` | horizon futur (12 par défaut) |
| `safety.max_deletions` | seuil de déclenchement du garde-fou |
| `mirror.subject_prefix` | préfixe des titres copiés ; `{label}` = agenda source |
| `mirror.hide_subject` | `true` : le titre réel n'est pas recopié, le préfixe devient tout le titre |
| `mirror.copy_location` / `copy_body` | passer à `false` pour ne pas recopier lieu et description |
| `calendars.<id>.backend` | `outlook_com`, `google_cal` ou `ics_readonly` |
| `calendars.<id>.store` | backend Outlook : nom du magasin, tel qu'affiché par `check_outlook.py` |
| `calendars.<id>.folder_path` | backend Outlook : sous-calendrier sous le Calendrier, ex. `Pro B` |
| `calendars.<id>.shared_with` | backend Outlook : adresse d'une personne dont l'agenda vous est partagé — alternative quand le compte ne peut pas être ajouté au profil |
| `calendars.<id>.writable` | `false` pour un agenda Outlook partagé sans droit de modification : il alimente les autres agendas sans jamais rien recevoir |
| `calendars.<id>.enabled` | `false` pour sortir temporairement un agenda du jeu |

Pour des copies discrètes entre les deux agendas professionnels :

```yaml
mirror:
  hide_subject: true
  subject_prefix: "Occupé — {label}"
  copy_location: false
  copy_body: false
```

Les copies n'affichent alors que « Occupé — Pro B » et le créneau : ni le titre
réel, ni le lieu, ni la description ne sortent de leur agenda d'origine.

---

## Couverture

| Direction | Statut |
|---|---|
| Pro A → Google | OK |
| Google → Pro A | OK |
| Pro B → Pro A | OK (titre et lieu seulement) |
| Pro B → Google | OK (titre et lieu seulement) |
| Pro A → Pro B | impossible |
| Google → Pro B | impossible |

**Pourquoi Pro B ne reçoit rien.** Écrire dans un agenda Microsoft 365 suppose soit une
application enregistrée dans le tenant (non autorisé par l'administration du tenant), soit le compte ajouté à un
profil Outlook local. Les trois voies ont été essayées :

1. *Ajout du compte Pro B au profil Outlook* — refusé.
2. *Partage de l'agenda Pro B vers votre adresse Pro A en « peut modifier »* — ce niveau
   n'est pas proposé par l'organisation, qui plafonne à « titres et lieux ». Et même accepté, un
   partage Microsoft 365 inter-organisations ne se matérialise pas en MAPI : Outlook
   classique crée un dossier vide et `GetSharedDefaultFolder` échoue.
3. *Publication d'un flux ICS depuis Outlook Web* — **retenu**, en lecture seule.

Pro B alimente donc les deux autres agendas sans jamais rien recevoir. Les descriptions
ne traversent pas : cette organisation ne publie que titres et lieux.

> Le flux ICS est régénéré périodiquement par Microsoft : un événement Pro B peut mettre
> plusieurs heures à apparaître ailleurs.

---

## Faire marche arrière

```bash
python sync.py --purge           # liste les copies qui seraient retirées
python sync.py --purge --apply   # les retire
```

La purge supprime **uniquement les copies** posées par le script, reconnues à leur
marqueur. Vos événements d'origine ne sont jamais touchés. C'est le filet qui permet
d'essayer la synchronisation sans s'engager : si le résultat ne convient pas, une purge
remet les trois agendas dans leur état initial.

---

## Exécution planifiée

Une tâche **« Synchro agendas »** est installée : du lundi au vendredi, toutes les heures
de 8 h à 19 h, via [run-sync.cmd](run-sync.cmd) qui journalise dans `logs/schedule.log`.

Deux contraintes à connaître :

- **Session Windows ouverte obligatoire.** La tâche pilote Outlook classique en COM, qui
  n'existe pas en session fermée. Elle est donc déclarée « Exécuter uniquement si
  l'utilisateur est connecté ».
- **Jeton Google.** Tant que l'écran de consentement OAuth reste en statut *Test*, Google
  l'invalide tous les 7 jours et la tâche échouera jusqu'à ce que vous relanciez
  `python -u authorize_google.py`. Le journal indique alors `autorisation Google expiree`.
  Publier l'application est la seule façon de rendre la planification vraiment autonome.

Pour la recréer ou l'ajuster, utiliser PowerShell plutôt que `schtasks` : ce dernier
interprète mal `/ET` sur un déclencheur hebdomadaire et produit une tâche sans prochaine
exécution.

```powershell
$action  = New-ScheduledTaskAction -Execute "C:\Projets\PERSOgendas
un-sync.cmd"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 08:00
$modele  = New-ScheduledTaskTrigger -Once -At 08:00 -RepetitionInterval (New-TimeSpan -Hours 1) -RepetitionDuration (New-TimeSpan -Hours 11)
$trigger.Repetition = $modele.Repetition
$settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 1)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "Synchro agendas" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force
```

Lancer, inspecter ou supprimer :

```powershell
Start-ScheduledTask    -TaskName "Synchro agendas"
Get-ScheduledTaskInfo  -TaskName "Synchro agendas"
Unregister-ScheduledTask -TaskName "Synchro agendas" -Confirm:$false
```

---

## Tests

Aucun accès à Outlook ni au réseau n'est nécessaire : les agendas sont simulés.

```bash
python tests/test_engine.py      # règles de réconciliation
python tests/test_roundtrip.py   # cycle complet, convergence et idempotence
```

---

## Dépannage

| Symptôme | Cause et remède |
|---|---|
| `magasin 'x' introuvable` | Le nom ne correspond pas à Outlook. Lancer `check_outlook.py` et recopier la valeur exacte. |
| `calendrier 'Pro A' introuvable` | Le message liste les calendriers Google existants ; corriger `calendar_name`. |
| Les appels Google expirent | Le proxy d'entreprise. `sync.py` duplique déjà `HTTP_PROXY` en minuscules pour httplib2 ; vérifier que la variable est bien définie dans la session. |
| `autorisation Google expiree ou revoquee` | Attendu en statut *Test* : le jeton ne vit que 7 jours. Relancer `python -u authorize_google.py`, ou publier l'application pour supprimer l'expiration. |
| `Erreur 403 : access_denied` à l'autorisation | Le compte Google n'est pas dans la liste des testeurs de l'écran de consentement, ou l'application n'est pas publiée. |
| Aucune page ne s'ouvre à l'autorisation | Normal hors terminal interactif. `authorize_google.py` affiche l'URL : la copier dans le navigateur. |
| `Arrêt de sécurité : N suppressions` | Lire le plan. Si les suppressions sont légitimes, relancer avec `--force`. |
| `Arrêt de sécurité : un agenda semble momentanément vide` | Outlook n'a pas fini de descendre la boîte aux lettres. Attendre que le dossier Calendrier se remplisse, puis relancer. |
| Copies Outlook décalées de l'offset horaire | Régression connue si le code revient à `Start`/`End` : `pywin32` convertit tout datetime en UTC et Outlook le relit comme heure locale. Les copies horaires doivent passer par `StartUTC`/`EndUTC`, les journées entières par une chaîne ISO locale. Voir `_set_utc` dans [backends/outlook_com.py](backends/outlook_com.py). |
| Un événement Pro B met des heures à apparaître | Propre au flux ICS publié, régénéré périodiquement par Microsoft. Seul le passage de `pro_b` en `outlook_com` supprime ce délai. |

Les journaux détaillés sont dans `logs/sync-AAAAMMJJ.log`.

---

## Licence

MIT. Projet personnel, fourni tel quel.
Pages du projet : <https://jerom138.github.io/synccal/>
