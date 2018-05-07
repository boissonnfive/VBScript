# VBScript

## Les bases du langage

### Imposer la déclaration des variables ###

    Option Explicit

### Gérer les exceptions soi-même ###

    ' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
    ' Doit être ajouté dans chaque routine
    On Error Resume Next


### Mettre en commentaires (') ###

Tout code placé entre un `'` et une fin de ligne est considéré comme un commentaire.

    ' Un commentaire d'introduction
    Err.Clear   ' Un commentaire de fin de ligne.
    ' If Err.Number <> 0 Then ' Un code mis en commentaires


### Déclaration et utilisation des variables (Dim et =) ###

    Option Explicit ' Force la déclaration des variables
    Dim maVariable
    maVariable = "Bruno" ' Une chaîne de caractères
    maVariable = 4       ' Un nombre
    maVariable = Now()   ' Date et heure du jour
    maVariable = 2+3*24  ' Affectation du résultat de 2+3*24 à maVariable


### Les noms des variables

- Ne sont pas sensibles à la casse.
- Ne peuvent comporter plus de 255 caractères.
- Doivent commencer par une lettre.
- Peuvent comporter des lettres (a-z, A-Z).
- Peuvent comporter des chiffres (0-9).
- Ne peuvent comporter de point ou d'espace.

*NOTE : Une variable peut être définie n'importe où dans le script car le compilateur lit le script en entier avant de l'exécuter.*


### Les tableaux ###

- Déclaration         : Dim Tableau(10) (déclare un tableau de 11 élément)
- Accès               : element = Tableau(3) (element prend la valeur du 4ème élément de tableau)
- Modification        : Tableau(0) = "coucou" (met coucou dans le 1er élément du tableau)
- Taille d'un tableau : nTaille = UBound(Tableau) (taille vaut 11)


### Les constantes ###

    Const ONE_HOUR = 3600000 ' Une heure en millisecondes


### Création et libération d'objets ###

    Option Explicit      ' Déclaration forcée des variables
    Dim objWShell        ' Déclaration de la variable qui va contenir l'objet
    Set objWShell = WScript.CreateObject("WScript.Shell") 
                         ' Affectation de la variable avec l'objet WScript.Shell
    Set objWShell = nothing 
                         ' Destruction de l'objet WshShell

*NOTE: on préfixera le nom d'un objet par obj ou o (ex: objWShell, oWShell)*


### Conditions (If/Then/Else/End If) ###

    If (expression1) Then
        instructions1
    ElseIf (expression2) Then
        instructions2
    Else
        instructions3
    End If


### Choix multiples (Select Case/End Select) ###

    Select Case expressiontest
        Case expression1 [, ...]  ' (ex : "3, 4, 5" => trois possibilité pour ce cas)
            instructions1
        
        Case expression2 [, ...] 
            instructions2
        
        Case Else
            instructions3
    End Select


### Boucles de répétitions ###

    For counter = start To end [Step step]
        [statements]
        [Exit For]
        [statements]
    Next

    For Each element In group
        [statements]
        [Exit For]
        [statements]
    Next [element]

    Do [{While | Until} condition]
       [statements]
       [Exit Do]
       [statements]
    Loop

    Do
       [statements]
       [Exit Do]
       [statements]
    Loop [{While | Until} condition]


### Créer et appeler une procédure ###

    Public Sub beep()
        call Wscript.CreateObject("wscript.Shell").Run("cmd /c @echo " & chr(7), 0)
    End Sub
    
    call beep() 'ou `beep`


### Créer et appeler une fonction ###

    Public Function dossierParent(nomCompletFichier)
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        dossierParent = objFSO.GetParentFolderName(nomCompletFichier)
        Set objFSO = Nothing
    End Function
    
    WScript.Echo "Dossier parent du script : " & dossierParent("WScript.ScriptFullName)



### Gérer des exceptions ###

    On Error Resume Next ' Nécessaire pour gérer soi-même les exceptions

    instructions
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'appel de la fonction OpenTextFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
        Err.Clear
    Else
        instructions
    End IF



### Renvoyer une valeur de retour ###

    DIM returnValue
    returnValue = 99
    WScript.Quit(returnValue)


### Les dates ###

**Date du jour**

    dateDuJour = Date()      ' =>  05/08/2016
    
**Heure du jour**

    heureDuJour = Time()     ' =>  14:34:07

**Date et heure du jour**

    dateHeureDuJour = Now() ' =>  05/08/2016   14:34:07

**Jour, numéro de jour, mois, année**

    aujourdHui = Now()
    WeekDayName(WeekDay(aujourdHui)) & " "_
    & WeekDay(aujourdHui) & " "_
    & MonthName(Month(aujourdHui)) & " "_
    & Year(aujourdHui)
    => Samedi 5 septembre 2015

**Date au format court**

    aujourdHui = Now()
    FormatDateTime(aujourdHui, vbShortDate) ' => 26/07/2016

**Date au format long**

    aujourdHui = Now()
    FormatDateTime(aujourdHui, vbLongDate) ' => mardi 27 juillet 2016

**Transforme une chaîne en date**

    aujourdHui = DateValue("26/07/2016")

**Compter le nombre de jours entre deux dates**

    nbJours = DateDiff("d", date1, date2) ' d pour compter les jours

**Date d'hier et de demain**

    dateHier = DateAdd("d",-1,Date) 'd: jour ; -1: un jour en moins; 
    dateDemain = DateAdd("d",1,Date) 'd: jour ; 1: un jour en plus; 


### Les fonctions prédéfinies utiles ###

- Space(NbSpaces)     : renvoie une chaîne de NbSpaces espaces
- vbNewLine ou vbCRLF : ajoute un saut de ligne dans une chaîne
- Sleep(Milliseconds) : pause le script le nombre de millisecondes spécifiées
- Timer               : Renvoie le temps écoulé depuis minuit (pour compter le temps d'exécution d'un script ou d'une fonction)


### Les chaînes de caractères (String) ###

- InStr(str1, str2)  : recherche str2 dans str1 et renvoie la position
- Split(str1, str2)  : découpe str1 en chaînes séparées par str2 (renvoie une collection)
- str = Replace(strInit, strFind, strReplace) : renvoie la chaîne remplacée
- UCase(str1)        : met str1 en majuscule
- LCase(str1)        : met str1 en minuscule

**Ajout de guillemets dans une chaîne de caractères :**

    strParameters = "mspaint.exe " & """" & JPGFileName & """"
    ou
    strParameters = "mspaint.exe " & chr(34) & JPGFileName & chr(34)



### Changer l'interpréteur de commandes ###

Interpréteur en ligne de commande par défaut (CScript.exe) :

    WSCRIPT //H:CScript
    CSCRIPT //H:CScript     

Interpréteur fenêtré par défaut (WScript.exe) :

    WSCRIPT //H:WScript
    CSCRIPT //H:WScript     

Cette commande n'est active que pour la session en cours. Si on veut la rendre définitive (jusqu'à un prochain changement), il faut ajouter le commutateur //S 

    WSCRIPT //H:CScript //S
    (Rend permanent l'utilisation de CSCRIPT comme interpréteur par défaut)


### Caractère de continuité de ligne ###

Pour une meilleure lisibilité, on peut être amené à écrire une instruction sur plusieurs lignes. Pour cela, on utilse le caractère de continuité de ligne "_".

    WScript.Echo "Je suis une phrase très longue qu'il faut _
    couper en deux, voir même en trois si on veut qu'elle _
    ne dépasse pas."


-------------------------------------------------------------------------------

## Cas pratiques ##

### Nom complet du script : WScript.ScriptFullName

### Répertoire courant (du script)

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(WScript.ScriptFullName)
    
    Wscript.Echo "Absolute path: " & objFSO.GetAbsolutePathName(objFile)
    Wscript.Echo "Parent folder: " & objFSO.GetParentFolderName(objFile)


### Parcourir le contenu d'un dossier ###

    set bubu = objFSO.GetFolder(objFSO.GetParentFolderName(WScript.ScriptFullName)).Files
    
    For each elt in bubu
        ' Wscript.echo objFSO.GetExtensionName(elt.Name)
        If objFSO.GetExtensionName(elt.Name) = "vbs" Then
            WScript.Echo elt.Name & "(" & elt.DateLastModified & ")"
        Else
            WScript.echo "..."
        End If
    Next


### Récupérer les arguments du script ###

    Set objArgs = WScript.Arguments
    Wscript.Echo objArgs.Count
    
    For Each objItem in objArgs
      Wscript.Echo objItem
    Next
    
    'For j=0 to objArgs.Count-1
     '   Wscript.Echo objArgs(j)
    'Next


### Lancer un logiciel avec des paramètres ###

    Dim WshShell
    Set WshShell = CreateObject("Wscript.Shell") 
    
    strParameters = "mspaint.exe " & """" & JPGFileName & """"
    ' Wscript.Echo strParameters
    
    WshShell.run strParameters
    'WshShell.Run("mspaint " & JPGFileName, 1, true)
    'Wscript.sleep 1000


### Utiliser des expressions régulières ###

    '------------------------------------------------------------------------
    ' Supprime toutes les balises HTML du texte en entr�e et renvoie le texte �pur�
    Function StripHTML(sTexteHTML)
        Dim oReg
        Set oReg = CreateObject("VBScript.RegExp")
        oReg.Pattern = "(<[^>]+>)"
        oReg.Global = True
        StripHTML = oReg.Replace(sTexteHTML, vbNullString)
    End Function


-------------------------------------------------------------------------------

## Pour aller plus loin

### Format ANSI des fichiers .vbs ###

Il faut enregistrer les fichiers .vbs au format ANSI, c'est-à-dire Windows1252 (ou codepage 1252) ou au format Unicode. Mais, pas au format UTF-8 qui n'est pas pris en charge par VBScript.

Pour pouvoir relire le fichier dans le bloc-notes, il vaut mieux utiliser des espaces à la place des tabulations.

Pour compiler dans Sublime Text 2, il faut utiliser le code page 850 (code page par défaut de l'invite de commande) dans le fichier de compilation du module VBScript :

    ...
    "windows":
    {
        "cmd": [ "cscript.exe", "/nologo", "$file" ],
        "encoding": "cp850"
    },
    ...


### WScript ou CScript ###

In addition to the two scripting languages, WSH provides two runtime programs: WScript.exe and CScript.exe. After you create a WSH script containing VBScript or Jscript commands, you use WScript or CScript to run the script. WScript runs the script as a Windows-based process and CScript runs the script as a console-based process.

source : <https://technet.microsoft.com/en-us/library/cc759559(v=ws.10).aspx>


### En-tête de script vbscript ###

    Option Explicit
    
    Const FICHIER       = "Creer_Dossiers_D.vbs"
    Const DESCRIPTION   = "Crée sur D: les dossiers nécessaires à l'installation    d'un ordinateur EXPANSIA."
    Const VERSION       = "3.3"
    Const AUTEUR        = "Bruno Boissonnet"
    Const DATE_CREATION = "22/07/2016"
    
    
    ' Remarques :
    ' - Les noms des dossiers sont dans la constante LISTE_DOSSIERS, séparés par    une virgule 
    ' - À enregistrer avec l'encodage ANSI
    ' - Utiliser "option explicit" pour forcer la déclaration des variables
    ' - Si on ne souhaite pas utiliser l'interface graphique :
    '     cscript.exe //NoLogo Creer_Dossiers_D.vbs > Creer_Dossiers_D.log
    
    
    ' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
    ' Doit être ajouté dans chaque routine
    ' On Error Resume Next


### Gestion des sorties des programmes vbscript ###

Il y a 3 modes d'affichages :

- mode fenêtré   : le programme est lancé à partir de l'explorateur windows et doit afficher des fenêtres.
- mode console   : le programme est lancé à partir de l'invite de commandes et doit afficher du texte.
- mode script    : le programme est lancé à partir d'une feuille HTA et ne doit rien afficher (il doit renvoyer un code d'erreur).


-------------------------------------------------------------------------------

## Visual Basic Script - Documentation ##

### Références ###

- Script Center : <https://technet.microsoft.com/en-us/library/bb902776.aspx>
- VBScript Reference : <https://technet.microsoft.com/en-us/library/ee198844.aspx>
- Programming with VBScript : <https://msdn.microsoft.com/en-us/library/aa227499%28v=vs.60%29.aspx?f=255&MSPPError=-2147217396>
- Scripting Guidelines : <https://technet.microsoft.com/en-us/library/ee198686.aspx>
- Scripting Concepts and Technologies for System Administration : <https://technet.microsoft.com/en-us/library/ee176762.aspx>
- Entreprise Script : <https://technet.microsoft.com/en-us/library/ee176576.aspx>


### WSH Object ("WScript") ###


- Windows Script Host Object Model : <https://msdn.microsoft.com/en-us/library/a74hyyw0%28v=vs.84%29.aspx>
- WSH Objects : <https://technet.microsoft.com/en-us/library/ee156581.aspx>
- Objet WScript : <https://msdn.microsoft.com/en-us/library/98591fh7%28v=vs.84%29.aspx>
    The WScript object is the root object of the Windows Script Host object model hierarchy. It never needs to be instantiated before invoking its properties and methods, and it is always available from any script file. The WScript object provides access to information such as: 
- Objet WScript.shell : <https://msdn.microsoft.com/en-us/library/aew9yb99%28v=vs.84%29.aspx>
    You create a WshShell object whenever you want to run a program locally, manipulate the contents of the registry, create a shortcut, or access a system folder. The WshShell object provides the Environment collection. This collection allows you to handle environmental variables (such as WINDIR, PATH, or PROMPT).
- Objet WScript.network : <https://msdn.microsoft.com/en-us/library/s6wt333f%28v=vs.84%29.aspx>


### Script Runtinme Object ("Scripting") ###

- Objet Scripting.FileSystemObject : <https://msdn.microsoft.com/en-us/library/6kxy1a51%28v=vs.84%29.aspx>
- Objet Scripting.Dictionnary : <https://msdn.microsoft.com/en-us/library/x4k5wbx4%28v=vs.84%29.aspx>


### Shell Object ("shell") ###

- Objet shell.application : <https://msdn.microsoft.com/en-us/library/windows/desktop/bb773938%28v=vs.85%29.aspx>
- InternetExplorer object : <https://msdn.microsoft.com/en-us/library/windows/desktop/aa752084%28v=vs.85%29.aspx>


### WMI Object ("winmgmts") ###

- <https://msdn.microsoft.com/en-us/library/aa394585%28v=vs.85%29.aspx>
 - Using WMI : <https://msdn.microsoft.com/en-us/library/aa393964%28v=vs.85%29.aspx>
- Creating a WMI Script : <https://msdn.microsoft.com/en-us/library/aa389763%28v=vs.85%29.aspx>


### Regular expression ###

- <https://msdn.microsoft.com/en-us/library/6wzad2b2%28v=vs.84%29.aspx>


### Divers ###

- How Can I Automatically Run a Script Any Time a File is Added to a Folder? : <http://blogs.technet.com/b/heyscriptingguy/archive/2004/10/11/how-can-i-automatically-run-a-script-any-time-a-file-is-added-to-a-folder.aspx>
- Fonctions sur chaînes de caractères (string) : <https://msdn.microsoft.com/fr-fr/library/e3s99sd8(v=vs.90).aspx>


-------------------------------------------------------------------------------

##  MODÈLE DE SCRIPT VBS


    '******************************************************************************
    '* Fichier     : CheminDossierParent.vbs                                      *
    '* Auteur      : Bruno Boissonnet                                             *
    '* Date        : 29/01/2015                                                   *
    '* Description : Renvoie le chemin complet du dossier parent d'un fichier.    *
    '*                                                                            *
    '* Remarques   :                                                              *
    '*               - Renvoie Empty s'il y a une erreur.                         *
    '******************************************************************************

    ' Force la d�claration des variables : on est oblig� de faire : `Dim Variable`
    Option Explicit

    ' Emp�che les erreurs de s'afficher (� supprimer lors du d�bogage)
    ' Doit �tre ajout� dans chaque routine
    On Error Resume Next

    Dim strFichier, strLine, result

    strFichier = "C:\Users\bubu.txt"
    strLine    = "Dossier parent du fichier """ & strFichier & """ :" & vbNewLine & CheminDossierParent (strFichier)
    result     = MsgBox(strLine, vbOKOnly+vbInformation, "CheminDossierParent.vbs")

    ' Message d'avertissement de fin de script.
    WScript.echo "Script termin�!"


    '******************************************************************************

    ' ***
    ' Nom              : CheminDossierParent.
    ' strCheminComplet : chemin complet du fichier.
    ' retour           : Le chemin du dossier parent termin� par un "\".
    ' ***
    Public Function CheminDossierParent(strCheminComplet)
    	On Error Resume Next
    	Dim objFSO, strCheminDossierParent, fin
    
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
    	strCheminDossierParent = objFSO.GetParentFolderName(strCheminComplet)
    	' Pas besoin de v�rification d'erreur car GetParentFolderName ne travaille
    	' pas sur des fichiers mais sur une cha�ne de caract�re.
    
    	Set objFSO = Nothing
    	' On ajoute une barre oblique invers�e au cas o� il n'y en aurait pas
    	fin = Right(strCheminDossierParent, 1)
    	if fin = "\" Then
    		CheminDossierParent = strCheminDossierParent
    	Else
    		CheminDossierParent = strCheminDossierParent  & "\" 
    	End If
    End Function




## A FAIRE ##

- [ ] Vérifier la présence d'un fichier /dossier
- [ ] Copier un fichier/dossier
- [ ] Supprimer un fichier/dossier
- [ ] Déplacer un fichier/dossier
- [ ] Renommer un fichier
- [ ] Récupérer le chemin du dossier contenant le script
- [ ] Vérifier si un dossier est vide
- [ ] Ouvrir un dossier dans l'explorateur Windows 
- [ ] Lancer un programme (avec des paramètres)
- [ ] Récupérer des valeurs de l'utilisateur (InputBox)
- [ ] Lire dans un fichier


A voir : <https://technet.microsoft.com/en-us/library/ee692824.aspx>
