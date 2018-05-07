' Scripts VBScript


' Le fichier de trace doit être fourni en argument pour indiquer si le script doit
' être lancé en console ou en graphique.

' Créer un script qui ajoute "TOut le monde" dans le groupe administrateur.


'------------------------------------------------------------------------------
' Sauts de ligne :
'------------------------------------------------------------------------------

' WScript.Echo "Saut de ligne : " & Chr(34)
' WScript.Echo "Saut de ligne : " & vbNewLine
' WScript.Echo "Saut de ligne : " & vbCRLF



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
' Parcours le contenu d'un dossier. Si les fichiers ont l'extension .vbs, 
' on affiche la date de dernière modification.
'------------------------------------------------------------------------------

Public sub DateDerniereModification()

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    set bubu = objFSO.GetFolder(objFSO.GetParentFolderName(WScript.ScriptFullName)).Files
    
    For each elt in bubu
    	' Wscript.echo objFSO.GetExtensionName(elt.Name)
    	If objFSO.GetExtensionName(elt.Name) = "vbs" Then
    		WScript.Echo elt.Name & "(" & elt.DateLastModified & ")"
    	Else
    		WScript.echo "..."
    	End If
    Next
End Sub


'------------------------------------------------------------------------------
' Parcours les disques de l'ordinateur et affiche des infos :
' Lettre, Nom, Espace libre, système de fichier
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

End Sub

'------------------------------------------------------------------------------
' Lit le contenu d'un fichier
' Applique un traitement sur le contenu
' Écrit dans un autre fichier
'------------------------------------------------------------------------------


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


Function LitFichier(cheminFichier)
Dim strFileContents
    ' Read the total content of the html file and put it in strFileContents
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTS = objFSO.OpenTextFile(cheminFichier)
    strFileContents = objTS.ReadAll
    objTS.Close
    LitFichier = strFileContents
End Function




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

' Émet le son d'alerte Windows par défaut (ce n'est pas un bip)
' Ne fonctionne que si le script est lancé par cscript'
Sub Bip
   WScript.Echo Chr(7)
End Sub

' Émet le son d'alerte Windows par défaut (ce n'est pas un bip)
Sub Biip
   CreateObject("WScript.Shell").Run "%comspec% /K echo " & Chr(07),0  'O: cache la fenêtre
End Sub


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
		    WScript.Echo "Erreur lors de la création de l'objet WScript.Shell." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
		    ' Err.Clear
    Else
		    strExplorerPath = "explorer.exe /e," & cheminDossier
    	  objShell.Run strExplorerPath
        Set objShell = Nothing
    End If
end sub


'------------------------------------------------------------------------------
' Nom           : AfficheDansExplorateur2
' Description   : Ouvre une fenêtre Explorateur Windows du dossier.
' cheminDossier : Chemin complet du dossier.
'------------------------------------------------------------------------------

public sub AfficheDansExplorateur2(cheminDossier)
   On Error Resume Next
   dim objShell
   set objShell = CreateObject("shell.application")
   if Err.Number <> 0 then
      WScript.Echo "Erreur lors de la création de l'objet shell.application." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
   else
      objShell.Explore(cheminDossier)
      set objShell = nothing
   end if
end sub


'------------------------------------------------------------------------------
' Nom           : Parle
' Description   : Parle le texte en paramètre.
' strTexte      : Texte à faire parler par l'ordinateur.
'------------------------------------------------------------------------------

Public Sub Parle(strTexte)
	CreateObject("SAPI.SpVoice").Speak strTexte
	' objVoice.Rate = 8 ' accélère le rythme du phrasé.
	' objVoice.Volume = 60
End Sub



' ***
' Nom      : DateDerniereModificationFichier
' filespec : Chemin complet du fichier
' retour   : Une date ou Empty s'il y a eu une erreur
' ***
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



'------------------------------------------------------------------------------
' Nom         : DisqueEstMonte
' Description : Dit si un disque est monté
' sLettreDisque  : Lettre correspondant au disque à tester
' retour      : Renvoie True si le disque est monté, False sinon
'------------------------------------------------------------------------------

Function DisqueEstMonte(sLettreDisque)

    Set FSys = CreateObject("Scripting.FileSystemObject")
    Set Drive = FSys.GetDrive(sLettreDisque & ":")
    DisqueEstMonte = Drive.IsReady

End Function


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
' Nom                 : Tracer
' Description         : Écrit dans un fichier le texte passé en paramètre.
' sCheminFichierTrace : Chemin du fichier de trace.
' sTexte              : texte à écrire dans le fichier de trace.
'------------------------------------------------------------------------------

' call Tracer("I:\vbs\adresses_mac.txt", "coucou")

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
' Nom                : RenommeFichier
' Description        : Renomme le fichier passé en paramètre
' sCheminFichier     : Chemin du fichier à renommer.
' sNomNouveauFichier : Chemin du fichier à renommer.
'------------------------------------------------------------------------------

Public Sub RenommeFichier(sCheminFichier, sNomNouveauFichier)

    Dim objFSO, sCheminCompletFichier
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    sCheminCompletFichier = DossierParent(sCheminFichier)
    objFSO.MoveFile sCheminFichier , sCheminCompletFichier & "\" & sNomNouveauFichier    ' renomme
    ' objFSO.MoveFile "D:\Bubu renommé.txt" , "D:\PERSONNEL\"  ' déplace

End Sub


### Renommer un dossier ###

    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.MoveFolder "D:\PERSONNEL" , "D:\PERSO"          ' renomme
    objFSO.MoveFolder "D:\PERSO" , "D:\INFORMATIQUE\"      ' déplace


### Vérifier qu un dossier existe ###

    strNomCompletDossier = "C:\Users\Bobo"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(strNomCompletDossier) Then
        WScript.echo "Le dossier n'existe pas"
        WScript.Quit
    End If


### Vérifier qu un fichier existe ###

    strNomFichier = "bubu.txt"
    strNomCompletDossier = "C:\Users\Bobo"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(strNomCompletDossier & "\" & strNomFichier) Then
        WScript.echo "Le fichier n'existe pas."
        WScript.Quit
    End If