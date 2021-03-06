'+----------------------------------------------------------------------------+
'| Fichier     : Fonction.vbs                                                 |
'+----------------------------------------------------------------------------+
'| Version     : 0.1                                                          |
'+----------------------------------------------------------------------------+
'| Description : Contient les fonctions que j'ai créées.                      |
'+----------------------------------------------------------------------------+






'+----------------------------------------------------------------------------+
'|                              FICHIERS                                      |
'+----------------------------------------------------------------------------+



'------------------------------------------------------------------------------
' Nom            : InfosFichier
' Description    : Affiche les infos du fichier (nom, extension, etc...)
' sCheminFichier : Chemin du fichier.
'------------------------------------------------------------------------------

Public Sub InfosFichier(sCheminFichier)

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(sCheminFichier)

    Wscript.Echo "Absolute path  : " & objFSO.GetAbsolutePathName(objFile)
    Wscript.Echo "Parent folder  : " & objFSO.GetParentFolderName(objFile) 
    Wscript.Echo "File name      : " & objFSO.GetFileName(objFile)
    Wscript.Echo "Base name      : " & objFSO.GetBaseName(objFile)
    Wscript.Echo "Extension name : " & objFSO.GetExtensionName(objFile)

End Sub


'------------------------------------------------------------------------------
' Nom                     : CreerFichier
' Description             : Crée un fichier texte dans le dossier précisé dans le chemin du fichier
' sCheminCompletFichier   : Chemin complet du fichier à créer.
'------------------------------------------------------------------------------

Sub CreerFichier(sCheminCompletFichier)

  On error resume next
  Dim objFSO, objFichierTest

  Set objFSO    = CreateObject("Scripting.FileSystemObject")
  set objFichierTest   = objFSO.CreateTextFile(sCheminCompletFichier)
  If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction CreateTextFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
  End If

End Sub


'------------------------------------------------------------------------------
' Nom                     : CopieFichier
' Description             : Copie un fichier dans le dossier spécifié
' sCheminCompletFichier   : Chemin complet du fichier à copier.
' sDossierDestination     : Chemin complet du dossier de destination.
' Remarques               : Ne pas oublier le "\" à la fin du chemin du dossier.
'------------------------------------------------------------------------------

Public Sub CopieFichier(sCheminCompletFichier, sDossierDestination)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CopyFile sCheminCompletFichier , sDossierDestination
    If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction CopieFichier." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
    End If
    Set objFSO = Nothing
End Sub


'------------------------------------------------------------------------------
' Nom                     : DeplaceFichier
' Description             : Déplace un fichier dans le dossier spécifié
' sCheminFichier          : Chemin complet du fichier à copier.
' sDossierDestination     : Chemin complet du dossier de destination.
' Remarques               : Ne pas oublier le "\" à la fin du chemin du dossier.
'------------------------------------------------------------------------------

Public Sub DeplaceFichier(sCheminFichier, sDossierDestination)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFile sCheminFichier , sDossierDestination
    If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction DeplaceFichier." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
    End If
    Set objFSO = Nothing
End Sub


'------------------------------------------------------------------------------
' Nom                : RenommeFichier
' Description        : Renomme le fichier passé en paramètre
' sCheminFichier     : Chemin du fichier à renommer.
' sNomNouveauFichier : Chemin du fichier à renommer.
'------------------------------------------------------------------------------

Public Sub RenommeFichier(sCheminFichier, sNomNouveauFichier)

    Dim objFSO, sCheminCompletFichier
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    sCheminCompletFichier = DossierParent(sCheminFichier)
    objFSO.MoveFile sCheminFichier , sCheminCompletFichier & "\" & sNomNouveauFichier

End Sub


'------------------------------------------------------------------------------
' Nom           : DateDerniereModificationFichier
' Description   : Renvoie la date de dernière modification du fichier (au format JJ/MM/AAAA)
' filespec      : Chemin complet du fichier
' Retour        : La date de dernière modification du fichier (au format JJ/MM/AAAA)
'------------------------------------------------------------------------------

Function DateDerniereModificationFichier(filespec)
   On Error Resume Next ' Emp�che les erreurs de s'afficher (� supprimer lors du d�bogage)
   Dim objFSO, objFile, retour, strErrMsg, result
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(filespec)
   If Err.Number <> 0 Then
      strErrMsg = "Erreur lors de l'appel de la fonction GetFile." & vbNewLine & "(Num�ro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
      result = MsgBox (strErrMsg, vbOKOnly+vbExclamation, "DateDerniereModificationFichier.vbs")
   Else
      retour = FormatDateTime(objFile.DateLastModified, 2) ' vbShortDate - 2 - Display a date using the short date format specified in your computer's regional settings.
   End If
   Set objFSO = Nothing
   Set objFile = Nothing
   DateDerniereModificationFichier = retour
End Function





'+----------------------------------------------------------------------------+
'|                              DOSSIERS                                      |
'+----------------------------------------------------------------------------+


'------------------------------------------------------------------------------
' Nom         : DossierParent
' Description : Renvoi le chemin du dossier parent
' cheminFichierOuDossier  : Chemin d'un fichier ou d'un dossier
' retour      : Le chemin du dossier parent de cheminFichierOuDossier
'------------------------------------------------------------------------------

Function DossierParent(cheminFichierOuDossier)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    DossierParent = objFSO.GetParentFolderName(cheminFichierOuDossier)
End Function


'------------------------------------------------------------------------------
' Nom            : TermineCheminParBarreOblique
' Description    : Ajoute une barre oblique à la fin du chemin si nécessaire.
' sCheminFichier : Chemin du fichier.
'------------------------------------------------------------------------------

Public Sub TermineCheminParBarreOblique(sCheminFichier)
  ' On ajoute une barre oblique inversée au cas où il n'y en aurait pas
  fin = Right(sCheminFichier, 1)
  if fin = "\" Then
    ' CheminDossierParent = strCheminDossierParent
  Else
    sCheminFichier = sCheminFichier  & "\" 
  End If
End Sub


'------------------------------------------------------------------------------
' Nom                     : CreeDossier
' Description             : Crée un dossier à partir du chemin contenu dans son nom.
' sCheminDossier          : Chemin complet du dossier à créer.
'------------------------------------------------------------------------------

Public Sub CreeDossier(sCheminDossier)
  Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder sCheminDossier
    If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction CreeDossier." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
    End If
    Set objFSO = Nothing
end sub


'------------------------------------------------------------------------------
' Nom            : RenommeDossier
' Description    : Renomme le dossier spécifié avec le nom spécifié.
' sCheminDossier : Chemin complet du dossier.
' sNouveauNom    : Nouveau nom du dossier.
'------------------------------------------------------------------------------

Sub RenommeDossier(sCheminDossier, sNouveauNom)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFolder sCheminDossier , objFSO.GetParentFolderName(sCheminDossier) & "\" & sNouveauNom
    set objFSO = Nothing
End sub


'------------------------------------------------------------------------------
' Nom            : DeplaceDossier
' Description    : Déplace le dossier spécifié dans le dossier destination spécifié.
' sCheminDossier : Chemin complet du dossier à déplacer.
' sCheminDossierDestinationAvecAntiSlash : Chemin complet du dossier destination.
'------------------------------------------------------------------------------

Sub DeplaceDossier(sCheminDossier, sCheminDossierDestinationAvecAntiSlash)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFolder sCheminDossier , sCheminDossierDestinationAvecAntiSlash
    set objFSO = Nothing
End sub


'------------------------------------------------------------------------------
' Nom            : DossierEstVide
' Description    : Dit si un dossier est vide ou pas
' sCheminDossier : Chemin complet du dossier
' retour         : Renvoie True si le dossier est vide, False sinon
'------------------------------------------------------------------------------

Public Function DossierEstVide(sCheminDossier)

    Dim objFSO, objFolder, retour
    retour = False
    'Set sCheminDossier = "C:\Documents and Settings\Marine Coite\Bureau\eMule\Temp"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(sCheminDossier)
    'Wscript.Echo objFolder.Size
    
    If objFolder.Size Then
        'Wscript.Echo "Le dossier " & strCheminCompletDossier & " n'est pas vide."
        retour = False
    Else
        'Wscript.Echo "Le dossier " & strCheminCompletDossier & " est vide."
        retour = True
    End If
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    DossierEstVide = retour
End Function


'------------------------------------------------------------------------------
' Nom            : MetALaCorbeille
' Description    : Mets l'élément spécifié à la corbeille.
' sCheminComplet : Chemin complet dde l'élément.
'------------------------------------------------------------------------------

Sub MetALaCorbeille(sCheminComplet)
  Dim oShellApp, dossier
  set oShellApp = WScript.CreateObject("Shell.Application")
  set dossier = oShellApp.Namespace(0).ParseName(sCheminComplet) 'Namespace(0) = Bureau
  dossier.InvokeVerb("delete")
  Set oShellApp = Nothing
end sub


'+----------------------------------------------------------------------------+
'|                               DISQUES                                      |
'+----------------------------------------------------------------------------+



'------------------------------------------------------------------------------
' Nom         : AfficheInfosDisques
' Description : Affiche les infos des disques de l'ordinateur
' Remarque    : Lettre, Nom, Espace libre, système de fichier
'------------------------------------------------------------------------------

Public sub AfficheInfosDisques()
    Set FSys = CreateObject("Scripting.FileSystemObject")
    Set AllDrives = FSys.Drives
    On Error Resume Next
    For Each iDrive In Alldrives
        s = s & "Lecteur " & iDrive.DriveLetter & " : - "
        s = s & iDrive.VolumeName & vbCrLf
        s = s & "Espace libre : " & FormatNumber(iDrive.FreeSpace/1024,  0)
        s = s & "Ko" & vbCrLf
        s = s & "System de fichier : " & iDrive.FileSystem
        s = s & vbCrLf
    Next
    MsgBox s
    Set FSys = Nothing
End Sub


'------------------------------------------------------------------------------
' Nom           : DisqueEstMonte
' Description   : Dit si un disque est monté
' sLettreDisque : Lettre correspondant au disque à tester
' retour        : Renvoie True si le disque est monté, False sinon
'------------------------------------------------------------------------------

Function DisqueEstMonte(sLettreDisque)
    Set FSys = CreateObject("Scripting.FileSystemObject")
    Set Drive = FSys.GetDrive(sLettreDisque & ":")
    DisqueEstMonte = Drive.IsReady
    Set FSys = Nothing
End Function




'+----------------------------------------------------------------------------+
'|                         LECTURE/ECRITURE                                   |
'+----------------------------------------------------------------------------+



'------------------------------------------------------------------------------
' Nom           : LitFichier
' Description   : Lit le contenu d'un fichier
' cheminFichier : le nom complet du fichier à lire
' retour        : Renvoie le contenu du fichier
'------------------------------------------------------------------------------

Function LitFichier(cheminFichier)
    Dim strFileContents
    ' Read the total content of the html file and put it in strFileContents
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTS = objFSO.OpenTextFile(cheminFichier)
    strFileContents = objTS.ReadAll
    objTS.Close
    LitFichier = strFileContents
End Function


Public sub SupprimeHTMLDuFichier(cheminFichier)

  Dim strFileContents
  Const ForAppending = 8
  Dim strNewFileName
  
  ' Read the total content of the html file and put it in strFileContents
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objTS = objFSO.OpenTextFile(cheminFichier)
  strFileContents = objTS.ReadAll
  objTS.Close
  
  ' Write the result of function in the file idem.txt
  Set objFile = objFSO.GetFile(cheminFichier)
  
  strNewFileName = objFSO.GetParentFolderName(objFile) & "\" & objFSO.GetBaseName(objFile) & ".txt"
  
  Set objTextFile = objFSO.OpenTextFile(strNewFileName, ForAppending, True)
  
  objTextFile.Write(StripHTML(strFileContents))
  
  objTextFile.Close

End Sub


'------------------------------------------------------------------------------
' Nom         : StripHTML
' Description : Supprime toutes les balises HTML du texte en entrée et renvoie le texte épuré
' sTexteHTML  : Texte HTML
' retour      : Renvoie le texte sans les balises HTML
'------------------------------------------------------------------------------

Function StripHTML(sTexteHTML)
  Dim oReg
    Set oReg = CreateObject("VBScript.RegExp")
    oReg.Pattern = "(<[^>]+>)"
    oReg.Global = True
    StripHTML = oReg.Replace(sTexteHTML, vbNullString)
End Function


'------------------------------------------------------------------------------
' Nom                 : Tracer
' Description         : Écrit dans un fichier le texte passé en paramètre.
' sCheminFichierTrace : Chemin du fichier de trace.
' sTexte              : texte à écrire dans le fichier de trace.
' Exemple             : call Tracer("I:\vbs\adresses_mac.txt", "coucou")
'------------------------------------------------------------------------------

Public Sub Tracer(sCheminFichierTrace, sTexte)

    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Set objFile = objFSO.CreateTextFile("I:\vbs\adresses_mac.txt")
    set objFile = objFSO.OpenTextFile(strCheminCompletFichierTrace, 8, True, -1) ' 8 = ForAppending, True pour créer le     fichier s'il n'existe pas, -1 pour écrire au format Unicode
    
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'appel de la fonction OpenTextFile (Numéro: " & Err.Number & ", Description: " &  Err.Description & ")"
        Err.Clear
    Else
        Dim MyVar
        MyVar = Now ' MyVar contains the current date and time.
        ' On écrit dans le fichier
        objFile.WriteLine MyVar & " " & strTrace
    
        ' On ferme le fichier
        objFile.Close
        Set objFile = Nothing
    End If
    
    Set objFSO = Nothing
End Sub



'------------------------------------------------------------------------------
' Nom              : Banniere
' Description      : Affiche le message en majuscule centré et encadré.
' sMessage         : Message à afficher.
' nLargeurBanniere : Largeur autorisée du texte.
'------------------------------------------------------------------------------


' +-------------------------------------------------------------------+
' |    BANNIERE.VBS - VERSION 0.1 - BRUNO BOISSONNET - 24/09/2015     |
' |        AFFICHE LE MESSAGE EN MAJUSCULE CENTRÉ ET ENCADRÉ.         |
' +-------------------------------------------------------------------+

Sub Banniere(ByVal sMessage, ByVal nLargeurBanniere)

    Dim sOutput, nEspace, sTrait, nSautDeLigne, sLigne, nNouvelleTailleDuMessage, nCompteur, nReste
    
    nCompteur = 0
    ' Création d'un trait de la forme : +--------- ... ---------+
    sTrait = "+" & String(nLargeurBanniere, "-") & "+"
    sOutput = vbCRLF & sTrait & vbCRLF

    ' On vérifie que la chaîne ne soit pas sur plusieurs ligne
    
    Do
        nSautDeLigne = InStr(sMessage, vbCRLF)
        ' WScript.Echo "nSautDeLigne = " & nSautDeLigne
        ' On vérifie que le message ne soit pas plus long que la largeur de la bannière
        If ( Len(sMessage) > nLargeurBanniere) Then
            ' WScript.Echo "La longueur du message est supérieur à la largeur de la bannière (" & Len(sMessage) & ">" & nLargeurBanniere & ")."

            ' On coupe le message en 2 lignes s'il n'y a pas déjà une ligne dedans
            If (nSautDeLigne = 0) Then
                sMessage = Left(sMessage, nLargeurBanniere - 1) & vbCRLF & Right(sMessage, Len(sMessage) - nLargeurBanniere + 1)
                ' WScript.echo "sMessage = " & sMessage
                nSautDeLigne = InStr(sMessage, vbCRLF)
                ' WScript.Echo "Nouveau nSautDeLigne = " & nSautDeLigne
            End If
        End If

        ' On vérifie que la première ligne du message ne soit pas plus longue que la largeur de la bannière
        
        If ( nSautDeLigne > nLargeurBanniere ) Then
            ' WScript.Echo "Erreur : la longueur d'une ligne du message est supérieure à la largeur de la bannière (" & nSautDeLigne & ">" & nLargeurBanniere & ")."
            sMessage = Left(sMessage, nLargeurBanniere - 1) & vbCRLF & Right(sMessage, Len(sMessage) - nLargeurBanniere + 1)
            ' WScript.echo "sMessage = " & sMessage
            nSautDeLigne = InStr(sMessage, vbCRLF)
            ' WScript.Echo "Nouveau nSautDeLigne = " & nSautDeLigne
            ' Exit Do
        End If


        ' Else
            ' Par défaut, on considère que le message n'est que sur une ligne
            sLigne = sMessage 
            ' Si un saut de ligne a été trouvé
            If nSautDeLigne <> 0 Then
                ' On récupère la chaîne avant le saut de ligne
                sLigne = Left(sMessage, nSautDeLigne - 1) ' -1 pour ne pas prendre le saut de ligne
                ' WScript.Echo "sLigne = " & sLigne
    
                ' On enlève cette chaîne et le saut de ligne du message
                nNouvelleTailleDuMessage = Len(sMessage) - nSautDeLigne - 1
                sMessage = Right(sMessage, nNouvelleTailleDuMessage)
                ' WScript.Echo "sMessage = " & sMessage
            End If

            nEspace = (nLargeurBanniere - Len(sLigne)) \ 2
            ' S'il y a un reste à la division, il faudra ajouter une espace à droite
            nReste = (nLargeurBanniere - Len(sLigne)) Mod 2
            sLigne = "|" & Space(nEspace) & sLigne & Space(nReste) & Space(nEspace) & "|" 
            sOutput = sOutput & UCase(sLigne) & vbCRLF
        ' End If
        nCompteur = nCompteur + 1
        If (nCompteur > 15) Then
            Exit Do
        End IF
    Loop While ( nSautDeLigne <> 0 )


    ' Centrage du message en ajoutant des espaces et des barres verticales :
    ' |     ... MESSAGE ...        |
    ' nEspace = (nLargeurBanniere - Len(sMessage)) \ 2
    ' sMessage = "|" & Space(nEspace) & sMessage & Space(nEspace) & "|" 

    ' sOutput = sOutput & UCase(sMessage) & vbCRLF
    sOutput = sOutput & sTrait & vbCRLF

    WScript.Echo(sOutput)

End Sub




'+----------------------------------------------------------------------------+
'|                              INTERNET                                      |
'+----------------------------------------------------------------------------+



'------------------------------------------------------------------------------
' Nom            : AdresseIP
' Description    : Renvoie l'adresse IP.
' retour         : Renvoie l'adresse IP de la carte connectée. "" sinon.
'------------------------------------------------------------------------------

public function AdresseIP()
   On Error Resume Next
   dim objWMIService, objColItems, objItem, strIP
   
   strIP = ""

   ' On crée un objet carte réseau
   'Set objNAC = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
   set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2") 
   set objColItems   = objWMIService.ExecQuery _ 
                       ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

   ' On parcours les propriétés de l'objet carte réseau
   for each objItem in objColItems
      ' On récupère l'adresse IP du PC
      if isNull(objItem.IPAddress) Then
         strIP = ""
      else
         strIP = objItem.IPAddress(0) 'adresse IPv4, IPv6 est dans (1)
      end if 

      exit for
    next

    AdresseIP = strIP
end function


'------------------------------------------------------------------------------
' Nom            : AdresseMAC
' Description    : Renvoie l'adresse MAC.
' retour         : Renvoie l'adresse MAC de la carte connectée. "" sinon.
'------------------------------------------------------------------------------

public function AdresseMAC()
   On Error Resume Next
   dim objWMIService, objColItems, objItem, strMAC
   
   strMAC = ""

   ' On crée un objet carte réseau
   'Set objNAC = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
   set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2") 
   set objColItems   = objWMIService.ExecQuery _ 
                       ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

   ' On parcours les propriétés de l'objet carte réseau
   for each objItem in objColItems
      ' On récupère l'adresse IP du PC
      if isNull(objItem.MACAddress) Then
         strMAC = ""
      else
         strMAC = objItem.MACAddress
      end if 

      exit for
    next

    AdresseMAC = strMAC
end function




'+----------------------------------------------------------------------------+
'|                             ORDINATEUR                                     |
'+----------------------------------------------------------------------------+




'------------------------------------------------------------------------------
' Nom         : ModeleOrdinateur
' Description : Renvoie le modèle de l'ordinateur
' retour      : Le modèle de l'ordinateur
'------------------------------------------------------------------------------

function ModeleOrdinateur()
  ' Déclaration des variables obligatoire
  Dim SystemName, objComputerSystem, ordinateur, retour
  SystemName = "localhost"
  
  set objComputerSystem = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
  SystemName & "\root\cimv2").InstancesOf ("Win32_ComputerSystem")
  for each ordinateur in objComputerSystem
    retour = trim(ordinateur.Manufacturer) & " " & trim(ordinateur.Model)
  Next
  
  Set objComputerSystem = Nothing
  Set ordinateur        = Nothing
  ModeleOrdinateur      = retour
   
end function


'------------------------------------------------------------------------------
' Nom         : NomOrdinateur
' Description : Renvoie le nom de l'ordinateur
' retour      : Le nom de l'ordinateur
'------------------------------------------------------------------------------

function NomOrdinateur() 
  On error resume next
  dim objWMIService, objColItems, objItem, strComputer, strNomOrdinateur 

  strComputer = "."

  set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
  set objColItems   = objWMIService.ExecQuery _ 
  ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True") 

  for each objItem in objColItems
    strNomOrdinateur = objItem.dnshostname
    exit for
  next

  NomOrdinateur = strNomOrdinateur

end function




'+----------------------------------------------------------------------------+
'|                            EXPLORATEUR                                     |
'+----------------------------------------------------------------------------+





'------------------------------------------------------------------------------
' Nom           : AfficheDansExplorateur
' Description   : Ouvre le dossier spécifié dans l'explorateur Windows.
' cheminDossier : Chemin complet du dossier.
'------------------------------------------------------------------------------

Public sub AfficheDansExplorateur(cheminDossier)
    ' On Error Resume Next
    Dim objShell, strExplorerPath
    Set objShell = CreateObject("Wscript.Shell")
    If Err.Number <> 0 Then
        ' WScript.Echo "Erreur lors de la création de l'objet WScript.Shell." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
        ' Err.Clear
    Else
    strExplorerPath = "explorer.exe /e," & cheminDossier
      objShell.Run strExplorerPath, , True
        Set objShell = Nothing
        WScript.Sleep 1500 ' Laisse le temps à la fenêtre de s'afficher
    End If
end sub

'------------------------------------------------------------------------------
' Nom           : AfficheDansExplorateur2
' Description   : Ouvre le dossier spécifié dans l'explorateur Windows.
' cheminDossier : Chemin complet du dossier.
'------------------------------------------------------------------------------

Public sub AfficheDansExplorateur2(cheminDossier)
    ' On Error Resume Next
    dim objShellAPP
    set objShellAPP = CreateObject("shell.application")
    objShellAPP.Explore(cheminDossier)
    set objShellAPP = Nothing
    WScript.Sleep 1500 ' Laisse le temps à la fenêtre de s'afficher
end sub


'------------------------------------------------------------------------------
' Nom           : CloseExplorerWindow
' Description   : Ferme le dossier spécifié dans l'explorateur Windows.
' sFolderPath   : Chemin complet du dossier.
'------------------------------------------------------------------------------

Function CloseExplorerWindow(sFolderPath)
    Dim bTest

    bTest = False
    
    with createobject("shell.application")
        
        ' WScript.Echo "Chemin en minuscule : " & lcase(sFolderPath)
        
        Do
            bSortieBoucle = false
            For Each wdow in .windows
                ' if typename(wndw.document) = "IShellFolderViewDual2" then
                ' WScript.Echo "Nombre de fenêtres : " & .windows.Count & vbCrLf & "Fenêtre en cours : " & wdow.document.folder.self.path
                
                If lcase(wdow.document.folder.self.path) = lcase(sFolderPath) Then
                  
                    WScript.Echo "Destruction de la fenêtre : " & wdow.document.folder.self.path
                    wdow.quit
                    bSortieBoucle = True
                    bTest = err.number = 0
                    
                     ' On quitte la boucle for si on a détruit une fenêtre
                     ' car il faut parcourir les fenêtres du début à nouveau
                     ' Pour vérifier s'il n'y en a pas une autre
                    Exit For
                end If
                
            Next
        Loop While bSortieBoucle = True
    
    end with ' shell.application
    CloseExplorerWindow = Cstr(bTest)
end function


'------------------------------------------------------------------------------
' Nom                : SelectionneFichierDansExplorateur
' Description        : Met le fichier en surbrillance dans l'explorateur.
' sNomCompletFichier : Chemin complet du fichier.
'------------------------------------------------------------------------------

Public Sub SelectionneFichierDansExplorateur(sNomCompletFichier)
    dim objWShell
    dim objShell
    dim objFolder
    dim objShellWindows
    dim sDossierParent
    dim sNomFichier


    set objWShell     = CreateObject("WScript.Shell")
    set objShell      = CreateObject("shell.application")
    set objFSO        = CreateObject("Scripting.FileSystemObject")
    
    sDossierParent    = objFSO.GetParentFolderName(sNomCompletFichier)
    sNomFichier       = objFSO.GetFileName(sNomCompletFichier)
    sNomDossierParent = objFSO.GetBaseName(sDossierParent)
    objFolder         = objShell.NameSpace(sDossierParent)
    

    ' WScript.Echo "Nom du fichier à sélectionner : " & vbCRLF & sNomFichier
    ' WScript.Echo "Nom du dossier parent : " & vbCRLF & sNomDossierParent

    
    ' WScript.Echo "Récupération de la fenêtre explorateur ..."
    
    set objShellWindows = objShell.Windows

    if (not objShellWindows is nothing) then
        
        ' WScript.Echo "Il y a " & objShellWindows.Count & " fenêtres ouvertes."
        For j=0 to objShellWindows.Count-1
          ' WScript.Echo "Nom de la fenêtre :" & vbCRLF & objShellWindows.Item(j).LocationName & vbCRLF & objShellWindows.Item(j).LocationURL

            If objShellWindows.Item(j).LocationName = sNomDossierParent Then

                ' WScript.Echo "Fenêtre trouvée !"

                set objShellFolderView = objShellWindows.Item(j).Document
                ' objShellFolderView.focus()
    
                For k=0 to objShellFolderView.Folder.Items.Count-1
    
                    ' WScript.Echo "Nom de l'élément :" & vbCRLF & objShellFolderView.Folder.Items.Item(k).Name

                    If objShellFolderView.Folder.Items.Item(k).Name = sNomFichier Then

                        objShellFolderView.SelectItem objShellFolderView.Folder.Items.Item(k),21
                        
                    End If
                Next

            End If
        Next

    end if

    set objWShell = nothing
    set objShell  = nothing
    set objFSO    = nothing

end Sub




'+----------------------------------------------------------------------------+
'|                               DIVERS                                       |
'+----------------------------------------------------------------------------+



'------------------------------------------------------------------------------
' Nom         : Beep
' Description : Emet un son d'erreur Windows.
'------------------------------------------------------------------------------

Public Sub Beep()
   On Error Resume Next
   Dim objVoice, objSpFileStream
   
   Set objVoice        = CreateObject("SAPI.SpVoice")
   Set objSpFileStream = CreateObject("SAPI.SpFileStream")
   
   objSpFileStream.Open "C:\Windows\Media\Windows Error.wav"
   objVoice.SpeakStream objSpFileStream
   objSpFileStream.Close

   Set objVoice        = Nothing
   Set objSpFileStream = Nothing
End Sub


'------------------------------------------------------------------------------
' Nom         : Bip
' Description : Émet le son d'alerte Windows par défaut (ce n'est pas un bip).
' Remarque    : Ne fonctionne que si le script est lancé par cscript
'------------------------------------------------------------------------------

Sub Bip
   WScript.Echo Chr(7)
End Sub


'------------------------------------------------------------------------------
' Nom         : Biip
' Description : Émet le son d'alerte Windows par défaut (ce n'est pas un bip).
'------------------------------------------------------------------------------

Sub Biip
   CreateObject("WScript.Shell").Run "%comspec% /K echo " & Chr(07),0  'O: cache la fenêtre
End Sub


'------------------------------------------------------------------------------
' Nom           : Parle
' Description   : Fait dire le texte en paramètre par l'ordinateur.
' strTexte      : Texte à faire parler par l'ordinateur.
'------------------------------------------------------------------------------

Public Sub Parle(strTexte)
  CreateObject("SAPI.SpVoice").Speak strTexte
  ' objVoice.Rate = 8 ' accélère le rythme du phrasé.
  ' objVoice.Volume = 60
End Sub






'+----------------------------------------------------------------------------+
'|                               TESTS                                        |
'+----------------------------------------------------------------------------+


' DateDerniereModification()

' AfficheInfosDisques()

' SupprimeHTMLDuFichier(cheminFichier)

' Dim contenu
' contenu = LitFichier("C:\Users\Bruno\Dropbox\En cours\vbscript\Test_hote_du_script.vbs")
' WScript.echo "Contenu du fichier Test_hote_du_script.vbs : " & vbCRLF & contenu

' Beep()
' Bip()
' Biip()
' AfficheDansExplorateur("C:\Users\Bruno\Dropbox\En cours\batch\Conversion_cp850")
' AfficheDansExplorateur2("C:\Users\Bruno\Dropbox\En cours\batch\Conversion_cp850")
' Parle("Bonjour Bruno. Tu es vraiment le meilleur !")

' Dim chemin
' chemin = DossierParent("C:\Users\Bruno\Dropbox\En cours\batch\Conversion_cp850")
' WScript.echo "Dossier parent de ""C:\Users\Bruno\Dropbox\En cours\batch\Conversion_cp850"" :" & vbCRLF & chemin 


' Dim objShell, strCheminBureau, strCheminFichierTest, strCheminDossierTest

' Set objShell         = CreateObject("Wscript.Shell")
' strCheminBureau      = objShell.SpecialFolders("Desktop")
' strCheminFichierTest = strCheminBureau & "\" & "test.txt"
' strCheminDossierTest = strCheminBureau & "\" & "Dossier test VBScript"

' CreerFichier strCheminFichierTest
' WScript.Sleep(2000)
' CreeDossier strCheminDossierTest
' WScript.Sleep(2000)
' CopieFichier strCheminFichierTest, strCheminDossierTest & "\"
' WScript.Sleep(2000)
' RenommeFichier strCheminFichierTest, "test2.txt"
' WScript.Sleep(2000)
' DeplaceFichier strCheminBureau & "\" & "test2.txt", strCheminDossierTest & "\"
' WScript.Sleep(5000)
' MetALaCorbeille(strCheminDossierTest & "\")




'+----------------------------------------------------------------------------+
'|                                 FIN                                        |
'+----------------------------------------------------------------------------+










'+----------------------------------------------------------------------------+
'|        FONCTIONS EXISTANTES RECODÉES PAR MOI POUR APPRENDRE                |
'+----------------------------------------------------------------------------+

' ----
' NomDossierScript
' Renvoie le nom complet du dossier contenant le script
' sans antislash à la fin
' ---
Function NomDossierScript

    Dim nLongueurNomDossier 
    
    ' Pour récupérer le dossier du script
    ' 1. On calcule la longueur de la chaîne représentant le nom du dossier
    ' C'est la taille totale moins la taille du nom du script
    nLongueurNomDossier = Len(WScript.ScriptFullName)  - Len(WScript.ScriptName)
    
    ' 2. On récupère cette longueur de chaîne dans le nom du script complet, en partant de la gauche (on enlève 1 pour ne pas prendre l'antislash)
    NomDossierScript = Left(WScript.ScriptFullName, nLongueurNomDossier -1 )

End Function


' ---
' NomDossierContenant
' Renvoie le nom complet du dossier contenant à partir du nom complet
' et du nom du fichier.
' sans antislash à la fin
' ---
Function NomDossierContenant(sNomComplet, sNom)

    Dim nLongueurNomDossier 
    
    ' Pour récupérer le dossier contenant
    ' 1. On calcule la longueur de la chaîne représentant le nom du dossier
    ' C'est la taille totale moins la taille du nom du fichier
    nLongueurNomDossier = Len(sNomComplet)  - Len(sNom)
    
    ' 2. On récupère cette longueur de chaîne dans le nom complet, en partant de la gauche (on enlève 1 pour ne pas prendre l'antislash)
    NomDossierContenant = Left(sNomComplet, nLongueurNomDossier -1 )

End Function


' ---
' NomFichier
' Renvoie le nom du fichier à partir de son nom complet.
' ---
Function NomFichier(sNomComplet)

    Dim nPositionDernierAntiSlash, nLongueurNomFichier 
    
    ' Pour récupérer le nom du fichier
    ' 1. On récupère la position du dernier antislash
    ' C'est la taille totale moins la taille du nom du fichier
    nPositionDernierAntiSlash  = InStrRev(sNomComplet, "\")
    
    ' 2. On calcule la longueur du nom du fichier à partir cette position
    nLongueurNomFichier = Len(sNomComplet) - (nPositionDernierAntiSlash)
    
    ' 3. On récupère la chaîne de cette longueur à partir de la droite
    NomFichier = Right(sNomComplet, nLongueurNomFichier)

End Function


