'===============================================
'WORDSTYLESINDEXOR
'https://github.com/arnauddalayer/WordStylesIndexor
'Arnaud d'Alayer
'Version : 20201218
'
'Cette cr�ation est mise � disposition selon le Contrat Paternit�-NonCommercial-ShareAlike2.5 Canada disponible en ligne http://creativecommons.org/licenses/by-nc-sa/2.5/ca/ ou par courrier postal � Creative Commons, 559 Nathan Abbott Way, Stanford, California 94305, USA.
'
'===============================================

Option Explicit
Dim stylesAIndexer : stylesAIndexer = Array("Titre 1", "Titre 2")

Dim separateur : separateur = "|" 'S�parateur de champs dans le fichier CSV
Dim listeChamps : listeChamps = "id" & separateur & "Fichier" & separateur & "Style" & separateur & "contenu" & vbCrLf 'ent�te du fichier CSV contenant la liste des champs
dim csv : csv = "rapport.csv"
dim db : db = "monIndex.accdb"
dim connStr : connStr = "provider=Microsoft.ACE.OLEDB.12.0; data source=" & db


'D�terminer le chemin actif
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
dim repertoireCourant : repertoireCourant = objFSO.GetAbsolutePathName(".")
csv = repertoireCourant & "\" & csv
db = repertoireCourant & "\" &  db



'===============================================
'ANALYSE DE LA COLLECTION ET CR�ATION D'UN R�SULTAT CSV
'===============================================
'SOURCE : http://www.microsoft.com/technet/scriptcenter/resources/qanda/may08/hey0529.mspx

Dim rapport 'Rapport CVS en m�moire avant �criture dans fichier
Dim nombreStylesAIndexer : nombreStylesAIndexer = UBound(stylesAIndexer)
Dim i : i = 0 'Compteur pour la liste des styles � parcourir dans chaque document
dim extraction 'Contient la cha�ne de caract�res extraite, pour faire quelques op�rations de nettoyage
dim j : j = 0 'Cl� primaire pour la BD

Dim objWord : Set objWord = CreateObject("Word.Application")
objWord.Visible = True

' On parcourt chaque fichier du dossier
Dim objFolder
Dim objFile
Dim strFilePath
Dim strExtension
Dim objDoc
Dim objSelection
Set objFolder = objFSO.GetFolder(repertoireCourant)
For Each objFile in objFolder.Files
	strFilePath = objFile.Path
	strExtension = objFSO.GetExtensionName(strFilePath)
	
	'Si le fichier est un document Word, on lance l'extraction
	If strExtension = "doc" Or strExtension = "docx" Then
		
		WScript.Echo "Traitement du fichier " & objFile
		
		'Parcours des styles
		for i = 0 to nombreStylesAIndexer
			WScript.Echo "    Traitement du style " & stylesAIndexer(i)
			Set objDoc = objWord.Documents.Open(strFilePath)
			Set objSelection = objWord.Selection
			
			objSelection.Find.ClearFormatting
			objSelection.Find.Forward = True
			objSelection.Find.Format = True
			
			'try/catch sur objSelection.Find.Style pour �viter erreur : "Microsoft Word: L'�l�ment dont le nom est sp�cifi� n'existe pas"
			On Error Resume Next
			Err.Clear
			objSelection.Find.Style = stylesAIndexer(i)
			If Err.Number <> 0 Then
				WScript.Echo "        Erreur : Le style " & stylesAIndexer(i) & " n'existe pas"
			Else
				While objSelection.Find.Execute
					If objSelection.Find.Found Then
						extraction = objSelection.Text
						'Suppression du caract�re du saut de ligne (13) et du saut de page (FF) http://fr.wikipedia.org/wiki/ASCII
						extraction = Replace(extraction, chr(12), "")
						extraction = Replace(extraction, chr(13), "")
						'V�rifier si ce n'est pas une chaine vide
						if len(extraction)>1 Then
							rapport = rapport & j & separateur & objFile.Name & separateur & stylesAIndexer(i) & separateur & extraction & separateur & vbCrLf
							j = j + 1
						end if
						'Correction documents Office 2010 : reprendre la recherche apr�s le dernier r�sultat pour �viter une boucle sans fin
						objSelection.Start = objSelection.End + 1
						objSelection.End = objSelection.Start
					end if
				Wend
			End if
			On Error Goto 0
			
			objDoc.Close
		Next
	
	End If
Next
objWord.Quit

'�criture du fichier CSV
dim objTextFile
Set objTextFile = objFSO.CreateTextFile(csv)
objTextFile.Write listeChamps
objTextFile.Write rapport
objTextFile.Close



'===============================================
'CREATION D'UNE BD ACCESS POUR RECEVOIR LE RESULTAT CSV
'===============================================
'SOURCE : http://database-programming.suite101.com/article.cfm/how_to_create_an_access_database_with_vbscript
'd�claration des types (http://msdn.microsoft.com/en-us/library/ms675318%28VS.85%29.aspx)
const adInteger = 3 'Integer
const adVarChar = 202 'Variable Character
const adLongVarChar = 203 'Memo

'S'il existe d�j� une base de donn�es, elle sera pr�alablement d�truite
if objFSO.FileExists(db) then
	objFSO.deletefile(db)
	'La date de cr�ation du nouveau fichier est la m�me que celle de l'ancien fichier. R�ponse : http://www.experts-exchange.com/Programming/Languages/.NET/Q_22589083.html (file system tunnelling)
end if

'V�rifier la pr�sence de ADOX, qui n'est plus expos� avec Office 365 (https://docs.microsoft.com/en-us/office/troubleshoot/access/cannot-use-odbc-or-oledb)
'Handling Errors
'SOURCE : https://stackoverflow.com/questions/4999364/try-catch-end-try-in-vbscript-doesnt-seem-to-work
On Error Resume Next
Err.Clear

dim catalog : set catalog = createobject("adox.catalog")
catalog.create connStr

If Err.Number <> 0 Then
	WScript.Echo "Impossible de creer une base de donnees Access. Veuillez installer Microsoft Access Database Engine 2016 Redistributable 32 bits (https://www.microsoft.com/en-us/download/details.aspx?id=54920) en mode /quiet"
	WScript.Quit
end If
On Error Goto 0

dim new_table : set new_table = createobject("adox.table")
new_table.Name = "monIndex"

new_table.columns.append "id", adInteger
new_table.columns.append "Fichier", adVarChar, 50
new_table.columns.append "Style", adVarChar, 50
new_table.columns.append "contenu", adLongVarChar
new_table.keys.append "id", 1, "id" 'unique id

catalog.Tables.Append new_table

set new_table = nothing
set catalog = nothing



'===============================================
' INSERTION DES DONN�ES DU CSV DANS BD ACCESS
'===============================================
'SOURCE : http://www.microsoft.com/technet/scriptcenter/resources/qanda/feb07/hey0206.mspx

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const ForReading = 1

Dim objConnection : Set objConnection = CreateObject("ADODB.Connection")
Dim objRecordSet : Set objRecordSet = CreateObject("ADODB.Recordset")
objConnection.Open connStr
objRecordSet.Open "SELECT * FROM monIndex", objConnection, adOpenStatic, adLockOptimistic

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(csv)

'Passer la 1�re ligne qui contient l'ent�te
dim ligneCSV : ligneCSV = objFile.ReadLine
Dim champsLigneCSV

Do Until objFile.AtEndOfStream
	ligneCSV  = objFile.ReadLine
	champsLigneCSV = Split(ligneCSV, separateur)

	objRecordSet.AddNew
	objRecordSet("ID") = CInt(champsLigneCSV(0))
	objRecordSet("Fichier") = champsLigneCSV(1)
	objRecordSet("Style") = champsLigneCSV(2)
	objRecordSet("Contenu") = champsLigneCSV(3)
	objRecordSet.Update

Loop

objRecordSet.Close
objConnection.Close