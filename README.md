# VBScript

Mes scripts VBScript


## Fichier vbscript.md

Contient mes notes sur le langage VBScript.


## Fichier Fonctions.vbs

Contient les fonctions que j'ai créées.

**FICHIERS**

- InfosFichier : Affiche les infos du fichier (nom, extension, etc...)
- CreerFichier : Crée un fichier texte dans le dossier précisé dans le chemin du fichier
- CopieFichier : Copie un fichier dans le dossier spécifié
- DeplaceFichier : Déplace un fichier dans le dossier spécifié
- RenommeFichier : Renomme le fichier passé en paramètre
- DateDerniereModificationFichier : Renvoie la date de dernière modification du fichier (au format JJ/MM/AAAA)
 

**DOSSIERS**
 
- DossierParent : renvoie le chemin du dossier parent
- TermineCheminParBarreOblique : Ajoute une barre oblique à la fin du chemin si nécessaire.
- CreeDossier : Crée un dossier à partir du chemin contenu dans son nom.
- RenommeDossier : Renomme le dossier spécifié avec le nom spécifié.
- DeplaceDossier : Déplace le dossier spécifié dans le dossier destination spécifié.
- DossierEstVide : Dit si un dossier est vide ou pas
- MetALaCorbeille :  Mets l'élément spécifié à la corbeille.


**DISQUES**

- AfficheInfosDisques : affiche les infos des disques de l'ordinateur
- DisqueEstMonte : Dit si un disque est monté


**LECTURE/ECRITURE**

- LitFichier : Lit le contenu d'un fichier
- SupprimeHTMLDuFichier : Supprime toutes les balises HTML du fichier
- Tracer : Écrit dans un fichier le texte passé en paramètre.
- Banniere : Affiche le message en majuscule centré et encadré.


**INTERNET**

- AdresseIP : Renvoie l'adresse IP.
- AdresseMAC : Renvoie l'adresse MAC.

**ORDINATEUR**

- ModeleOrdinateur : Renvoie le modèle de l'ordinateur
- NomOrdinateur : Renvoie le nom de l'ordinateur


**EXPLORATEUR**

- AfficheDansExplorateur : Ouvre le dossier spécifié dans l'explorateur Windows.
- AfficheDansExplorateur2 : Ouvre une fenêtre Explorateur Windows du dossier.
- CloseExplorerWindow : Ferme le dossier spécifié dans l'explorateur Windows.
- SelectionneFichierDansExplorateur : Met le fichier en surbrillance dans l'explorateur.


**DIVERS**

- Beep : Emet un son d'erreur Windows.
- Bip : Émet le son d'alerte Windows par défaut (ce n'est pas un bip)
- Biip : Émet le son d'alerte Windows par défaut (ce n'est pas un bip).
- Parle : Fait dire le texte en paramètre par l'ordinateur.


**À faire :**

- Vérifier qu'un dossier existe
- Vérifier qu'un fichier existe