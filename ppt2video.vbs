' Write each slide's notes to a text file
' in same directory as presentation itself
' Each file is named NNNN_Notes_Slide_xxx
' where NNNN is the name of the presentation
'       xxx is the slide number
'PPT���ŵ�ʱ�䲻Ҫ������ʮ����
'PPT���������л�Ч��������ǳ��....
Option Explicit
Dim oSlide
Dim oShape
Dim strFileName
Dim strNotesText
Dim intFileNum
Dim inputFile
Dim objPPT
Dim objPresentation
Dim objFso
Dim objPres
Dim FileType
Dim WshShell
Dim Path
Dim ppPlaceholderBody
Dim strFilepath6
Dim strFilepath7 
Dim OpenClipName, FinalClipName
ppPlaceholderBody = 2  '���� Body
Dim dub_volume         '��������
Dim background_volume  '������������
Dim CurrentPath
dub_volume = 1
background_volume = -53

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

Set WshShell = WScript.CreateObject("WScript.Shell") 

Path = WshShell.CurrentDirectory    '��ȡ�ű�·��
WriteLine "�ű���·��Ϊ " & Path 

CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
WriteLine "��ǰ����Ŀ¼�ǣ�" & CurrentPath

If WScript.Arguments.Count < 1 Then
	WriteLine "PPT to Video v2.0"
	WriteLine "cscript ppt2video PPT/PPTx [t/T]"
	WriteLine "ʹ��˵����You need to specify input PPT/PPTx file��[t/T]ʹ��t����T����ʾ�����ǲ���ϳ���Ƶ�������к�����ת����Ƶ����"
	WScript.Quit
End If

'############################################################################################################################################
inputFile = CurrentPath & "\\pptx\\" & WScript.Arguments(0)	'Ҫ��PPT�ļ��ĵ�ַ
strFilepath6  = CurrentPath & "\\" & "MP4"				'���MP4���ļ���·��
'############################################################################################################################################

Set objFso = CreateObject("Scripting.FileSystemObject")

If objFso.FileExists( inputFile ) = False Then   '����PPT�Ƿ����
	WriteLine "�Ҳ���PPT/PPTx�ļ���" & inputFile
	WScript.Quit
End If

Set objPPT = CreateObject( "PowerPoint.Application" )

objPPT.Visible = True

Set objPres = objPPT.Presentations.Open(inputFile)

Set objPresentation = objPPT.ActivePresentation 


Dim strFilepath1,strFilepath2,strFilepath3,strFilepath4,strFilepath5

Dim strSrcFld, strDstFile

strFilepath1 = Path & "\\"  
strFilepath7  = strFilepath6 & "\\"& objPresentation.Name & ".mp4"      '�����MP4���ļ�·��                                 

'������ĵ�ַ��������Ķ�������Ҫ��ffmpeg������ִ�У���Ҫ��ffmpeg��װ������ͬһ���ļ�����

OpenClipName  = Path & "\\" & "open.mp4"                      'Ƭͷ·��
FinalClipName = Path & "\\" & "final.mp4"                     'Ƭβ·�� 
strFilepath2 =  Path & "\\" & "wav" & "\\"                    'PPT���ڵ��ļ�·����wav�ļ����µ��ļ���e:\PPT2Video\wav\��
strFilepath3 = strFilepath2 & "*.wav"                         'wav�ļ��������е�.wav�ļ���e:\PPT2Video\wav\*.wav��
strFilepath4  = strFilepath1 & "back" & "\\" & "backg.wav"    '�������ֵ�·��
strFilepath5  = Path  & "\\" & objPresentation.Name & ".mp4"
strSrcFld = strFilepath2                                      'Ҫ��ȡ.wav�ļ�����·��
strDstFile = strFilepath1 & "list.txt"                        '��ȡ�ļ�������list.txt��·��

' One --------------------------------------------------------------- ��һ������ȡPPT�����֣������ļ�

WriteLine "****************��ȡPPT������****************"

'ɾ��strFilepath1�е�����.txt��.wav�ļ���strFilepath2�е�����.wav�ļ�

	if objFso.fileexists(strFilepath1 & "1.txt")=True then
		objFso.deleteFile strFilepath1 & "???.txt"
		WriteLine "���е�txt�����ļ�ɾ���ɹ�........."
	End If
	
	if objFso.fileexists(strFilepath1 & "001.wav")=True then
		objFso.deleteFile strFilepath1 & "*.wav"
	End If
	
	if objFso.fileexists(strFilepath2 & "001.wav")=True then 
		objFso.deleteFile strFilepath2 & "*.wav"
		WriteLine "���е�wav��Ƶ�ļ�ɾ���ɹ�........"
	End If

'��ȡ���֣�����.txt�ļ�

Dim SlideNumber
'Dim cc
SlideNumber = 1

	For Each oSlide In objPresentation.Slides
		WriteLine("��ȡPPT��ҳ����" & "��" & SlideNumber & "ҳ" & "........")	
	
		For Each oShape In oSlide.NotesPage.Shapes
			If oShape.PlaceholderFormat.Type = ppPlaceholderBody Then
				If oShape.HasTextFrame Then
					strFileName = Path _
                        & "\" & CStr(oSlide.SlideIndex) _
                        & ".TXT"
					Set intFileNum=objFso.CreateTextFile(strFileName, True) 
					'cc=oShape.TextFrame.TextRange.Text
					'WriteLine cc
					intFileNum.Write(oShape.TextFrame.TextRange.Text)
					intFileNum.close 
				End If
			End If
		Next

		SlideNumber = SlideNumber + 1
	Next

' Two ---------------------------------------------------------------�ڶ���������ת����

WriteLine "****************����ת����****************"

If objFso.FolderExists(Path & "\\" & "wav")=false Then
	objFso.createfolder Path & "\\" & "wav"
	WriteLine "wav�ļ��д����ɹ�........"
End if

Dim objShell

Dim objWMIService, colProcessList, strComputer

strComputer = "."

Set objShell = CreateObject("Wscript.Shell")

objShell.Run "tts.exe" '����tts.exe����������ת��

Do While True
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'tts.exe'")
	IF colProcessList.Count > 0 THEN
	Else
		Exit Do	
	END IF
Loop

' Three ---------------------------------------------------------------���������ı�PPT�л�����ʱ��

WriteLine "****************�ı�PPT�л�����ʱ��****************"

Dim SlideNumber1       
Dim oExec, Return_Value
Dim SlideAudioFile
Dim msoAnimEffectMediaPlay, msoAnimTriggerAfterPrevious, msoAnimTriggerWithPrevious
msoAnimEffectMediaPlay = 83
msoAnimTriggerAfterPrevious = 3
msoAnimTriggerWithPrevious = 2

SlideNumber = 1
If objFso.FileExists( strFilepath1& "001.wav")=false Then
	objFso.CopyFile strFilepath3,strFilepath1,False
End If
	
'��ȡҳ����ҳ��������9ʱ������Ϊ001,002.......����ҳ������9��������Ϊ010,011........��
'Ϊ�˷����ȡ.wav�ļ���ʱ������˳�����У������������ӣ�����������	

Dim oShp, oEffect
For Each oSlide In objPresentation.Slides
	If  SlideNumber <=9 Then
		SlideNumber1 = "00" &SlideNumber
	Else
		SlideNumber1 = "0" &SlideNumber
	End If
	If objFso.FileExists( strFilepath1 & SlideNumber1 & ".wav" ) = False Then
		WriteLine "No notes, skip slide " & SlideNumber1
	Else
		SlideAudioFile = strFilepath1 & SlideNumber1 & ".wav"
		Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffprobe -loglevel error -show_streams " & SlideNumber1 & ".wav | grep duration= | cut -f2 -d=") 
		'��ȡ��Ƶʱ��
		WScript.echo oExec.StdErr.ReadAll()
		Return_Value = oExec.StdOut.ReadLine
		' add by lujun ���Ѷ�ɷ��ش���Ҳ����TTS���ɵ�wav�ļ��޷�������Ч���ȣ���ôȱʡ��ÿҳ���л�ʱ������Ϊ 1�����������������˳���
		If Return_Value = "N/A" Then
			Return_Value = 1
		End If
		WriteLine "��ȡPPT�л���ʱ��(Slide Num " & SlideNumber & "): " & Return_Value	& "��........"
		oSlide.SlideShowTransition.AdvanceOnClick = False
		oSlide.SlideShowTransition.AdvanceOnTime = True
		'oSlide.SlideShowTransition.Duration = 2
		'---------------------------------------------------------------------------------------	
		oSlide.SlideShowTransition.AdvanceTime = Return_Value			' �޸�ÿҳ�ĳ�������ʱ��
		'---------------------------------------------------------------------------------------			
		
		'ɾ��ԭ�����ӵ���������
		Dim intShape
		With oSlide.Shapes
			For intShape = .Count To 1 Step -1
				With .Item(intShape)
					If .Type = 16 And .Name = "PPT2Video" Then .Delete	' 16 ����ʹ�� Shapes.AddMediaObject2 ���ӵ���������Name�ж���Ϊ�˷�ֹ��ɾ����������������
				End With
			Next
		End With

		If WScript.Arguments.Count = 2 Then
			If WScript.Arguments(1) = "T" Or WScript.Arguments(1) = "t" Then
				Set oShp = oSlide.Shapes.AddMediaObject2(SlideAudioFile, True, False, 10, 10)
				oShp.Name = "PPT2Video"
				Set oEffect = oSlide.TimeLine.MainSequence.AddEffect(oShp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
				oEffect.MoveTo 1
				With oEffect
					.EffectInformation.PlaySettings.HideWhileNotPlaying = True
				End With
				With oShp.AnimationSettings.PlaySettings
					.PlayOnEntry = True
					.PauseAnimation = False
					.StopAfterSlides = 9999	'�����������붯��һ�𲥷�
				End With

								
			End if
		End If

	End If	
	SlideNumber = SlideNumber + 1
Next

objPresentation.Save

If WScript.Arguments.Count = 2 Then
	If WScript.Arguments(1) = "T" Or WScript.Arguments(1) = "t" Then
		objPresentation.Close
		ObjPPT.Quit
		WriteLine  ""
		WriteLine  "PPT����Ӻϳɵķ�����Ƶ�ɹ�............"
		WScript.Quit
	End if
End If
'four---------------------------------------------------------------���Ĳ�����PPTתΪMP4

WriteLine "****************PPTתΪ��Ƶ****************"

'objPresentation.CreateVideo(fileName, DefaultSlideDuration:=1, VertResolution:=480)
'���Թ�����pdf objPresentation.SaveAs "c:\\lujun.mp4", 32, False

'ע�⣡��������MP4�ļ�Ҫ�ķ��൱���ʱ�䣬��˺���� objPresentation.Close �� ObjPPT.Quit��������ִ�У�
'����ᵼ������MP4�ļ�������ȷ�������ǵȴ�MP4������Ϻ󣬲�ִ�� objPresentation.Close �� ObjPPT.Quit��

Dim ppSaveAsMP4,f,MP4File
ppSaveAsMP4 = 39   '39������mp4��ʽ����Ĳ���
MP4File = Path & "\" & objPresentation.Name & ".save.mp4"	'���ɵ�mp4·��

If objFso.FileExists( MP4File ) Then
	objFso.deleteFile MP4File
End If

If WScript.Arguments.Count = 2 Then
	'CreateVideo( FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality )
	'UseTimingsAndNarrations:=True, VertResolution:=720, FramesPerSecond:=30, Quality:=100
	'WScript.Arguments(1)��120��240��480��
	objPresentation.CreateVideo MP4File,True,3,WScript.Arguments(1),30,60
	'ppMediaTaskStatusNone				0
	'ppMediaTaskStatusInProgress		1
	'ppMediaTaskStatusQueued			2
	'ppMediaTaskStatusDone				3
	'ppMediaTaskStatusFailed			4
	'�ȴ�PPTת��ΪMP4���
	Do While True
		wscript.sleep 1000
		If objPresentation.CreateVideoStatus = 3 Or objPresentation.CreateVideoStatus = 4 Then
			Exit Do
		End if
	Loop
Else
	'------��������ת�����룬ת������Ƶ�����ߣ����Ǻ�ʱ�ܳ�
	objPresentation.SaveAs MP4File , ppSaveAsMP4, False			'��pptתΪ��Ƶ������mp4��ʽ����
End If

'�ȴ�PPTת��ΪMP4���
Do While True	
	If objFso.FileExists( MP4File ) = False Then			'�������Ҫwscript.sleep����Ȼ�Ҳ�����Ƶ��·��
		wscript.sleep 3000
	Else
		wscript.sleep 3000
		Set f = objFso.GetFile(MP4File)  '��ȡ��Ƶ�Ĵ�С
		If f.Size > 0 Then
			WriteLine MP4File & " �ļ���СΪ :" & f.Size & " Bytes........"
			Exit Do
		End if
	End If
Loop

Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & MP4File & " -vn -y -acodec copy video.aac")'����Ƶ����Ƶ�з������
WScript.echo oExec.StdErr.ReadAll()
If objFso.FileExists( strFilepath1 & "video.aac" ) Then
	WriteLine "��Ƶ����Ƶ�з���ɹ�"
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i video.aac video.wav") '����Ƶ��ʽתΪ.wav
	WScript.echo oExec.StdErr.ReadAll()
	objFso.deleteFile strFilepath1 & "video.aac"
	If objFso.FileExists( strFilepath1 & "video.aac" ) = False Then
		WriteLine "video.aac�ļ�ɾ���ɹ�"
	Else
		WriteLine "video.aac�ļ�ɾ��ʧ��"
	end if
Else
	WriteLine "��Ƶ����Ƶ�з���ʧ�ܣ�video.aac�ļ�������"
End If

If objFso.FileExists( strFilepath1 & "video.wav" ) = False Then
	WriteLine "video.wav�ļ�������"
Else
	WriteLine "video.wav�ļ�����"
End If
'five...................................................................���岽�������ɵ���������������������

WriteLine "****************����������������****************"

'��ȡwav�ļ�����.wav�ļ����ļ�������file '1.wav'��ʽ������list.txt�ļ���

Dim objFso1, objSrcFls, objFile, objDstFile

Set objFso1 = CreateObject("Scripting.FileSystemObject")

If objFso1.FileExists( strDstFile ) Then
	objFso.deleteFile strDstFile
End If

Set objDstFile = objFso1.OpenTextFile(strDstFile, 2, True)

Set objSrcFls = objFso1.GetFolder(strSrcFld).Files

For Each objFile In objSrcFls
	objDstFile.WriteLine " file " & "'" & objFile.Name & "'"
Next

Dim objFld, objSrcFld

Set objSrcFld = objFso1.GetFolder(strSrcFld).SubFolders

For Each objFld In objSrcFld
	Call LoopSubFlds(strSrcFld & objFld.Name & "\")
Next

objDstFile.Close

Set objFile = Nothing

Set objSrcFls = Nothing

Set objFso1 = Nothing

Sub LoopSubFlds(strFld)
	For Each objFile In objFso.GetFolder(strFld).Files
		ojDstFile.WriteLine " file " & "'" & objFile.Name & "'"
	Next
	For Each objFld In objFso.GetFolder(strFld).SubFolders
		Call LoopSubFlds(strFld & objFld.Name & "\")
	Next
End Sub

If objFso.FileExists(strFilepath1 & "notes.wav")=false Then
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -f concat -i list.txt -c copy notes.wav") '�������е�.wav�ļ���������Ƶ����
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "�������ɳɹ�........"
Else
	WriteLine "����������........"
End If 
	
'six...................................................................���������������뱳�����ֺϲ�	

WriteLine "****************�����뱳�����ֺϲ�****************"

If objFso.FileExists(strFilepath1 & "dnotes.wav")=false Then
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i notes.wav -af volume=" & dub_volume & "dB dnotes.wav")  '������������
	WScript.echo oExec.StdErr.ReadAll()
End If
	
If objFso.FileExists(strFilepath1& "backg.wav")=false Then
	objFso.CopyFile strFilepath4,strFilepath1,False
End If

Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i backg.wav -af volume=" & background_volume & "dB dbackg.wav")  '���ӱ�����������

WScript.echo oExec.StdErr.ReadAll()

if objFso.fileexists(strFilepath1 & "allvoice.mp3")=False then 
	WriteLine "�����ͱ������ֺϳ��ļ�������........"
else
	objFso.deleteFile strFilepath1 & "*.mp3"
	WriteLine "�����ͱ������ֺϳ��ļ�ɾ���ɹ�........"
End If
Dim v
If objFso.FileExists( strFilepath1 & "video.wav" ) = true Then
	WriteLine "video.wav�ļ�����........"
	Set v = objFso.GetFile(strFilepath1 & "video.wav")
	WriteLine "�������Ƶ�ļ���СΪ :" & v.Size & " Bytes........"
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i dnotes.wav -i dbackg.wav -i video.wav -filter_complex amix=inputs=3:duration=first:dropout_transition=2 -f mp3 allvoice.mp3")
	'���������������ĺϲ�
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "�������Ƶ�뱳���������ϳɳɹ�"
else
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i dnotes.wav -i dbackg.wav -filter_complex amix=inputs=2:duration=first:dropout_transition=2 -f mp3 allvoice.mp3")
	'���������������ĺϲ�
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "video.wav�ļ�������........"
	WriteLine "�����ͱ������ֺϳɳɹ�........"
End if
'seven...................................................................���߲����������������������ɵ���Ƶ����Ƶ�ϲ�

WriteLine "****************��Ƶ����Ƶ�ϲ�****************"

If objFso.FileExists(strFilepath1 & "video.mp4") Then
	objFso.deleteFile strFilepath1 & "video.mp4"
	WriteLine "video.mp4ɾ���ɹ�"
else
	WriteLine "video.mp4�ļ�������"
End If
If objFso.FileExists( strFilepath1 & "video.wav" ) = true Then
	WriteLine "video.wav�ļ�����"
	If objFso.FileExists(strFilepath1 & "partvideo.mp4")=true Then
		objFso.deleteFile strFilepath1 & "partvideo.mp4"
	else
		WriteLine "partvideo.mp4�ļ�������"
	end if
	If objFso.FileExists(strFilepath1 & "partvideo.mp4")=false Then
		If objFso.FileExists(MP4File) Then
			WriteLine MP4File & "����"
			Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i " & MP4File & " -an partvideo.mp4") '����Ƶ����ȡ��Ƶ����Ҫ��Ƶ
			WScript.echo oExec.StdErr.ReadAll()
			If objFso.FileExists( strFilepath1 & "partvideo.mp4" )Then
				WriteLine "partvideo.mp4���ɳɹ�"
				If objFso.FileExists( strFilepath1 & "video.mp4" ) = False Then
					Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i partvideo.mp4 -i allvoice.mp3 -r 25 video.mp4") '��Ƶ����Ƶ�ĺϲ�
					WScript.echo oExec.StdErr.ReadAll()
					objFso.deleteFile strFilepath1 & "partvideo.mp4"
					if objFso.FileExists(strFilepath1 & "partvideo.mp4")=false Then
						WriteLine "partvideo.mp4ɾ���ɹ�"
					else
						WriteLine "partvideo.mp4ɾ��ʧ��"
					end if
				else 
					WriteLine "video.mp4�ļ�����"
				end if 
			else
				WriteLine "partvideo.mp4����ʧ��"
			end if
		else
			WriteLine MP4File & "������"
		end if
	else
		WriteLine "partvideo.mp4�ļ�����"
	end if
else
	if objFso.FileExists(strFilepath1 & "video.mp4")=false Then
		Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i " & MP4File & " -i allvoice.mp3 -r 25 video.mp4") '��Ƶ����Ƶ�ĺϲ�
		WScript.echo oExec.StdErr.ReadAll()
		WriteLine "�������֣���������Ƶ�ϳɳɹ�........"
	Else 
		WriteLine "video.mp4�ļ�������"
	End if 
End if


'eight.....................................................................�ڰ˲�:Ƭͷ��Ƭβ����Ƶ�ϲ�	

WriteLine "*****************Ƭͷ��Ƭβ����Ƶ�ϲ�****************"

If objFso.FileExists(strFilepath1 & "video.ts")=true Then
	objFso.deleteFile strFilepath1 & "video.ts"
End If
	
Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i video.mp4 -c copy -bsf:v h264_mp4toannexb -f mpegts video.ts") '����Ƶ����һ��.ts��ʽ��������Ƶ������

WScript.echo oExec.StdErr.ReadAll()

If objFso.FileExists(OpenClipName)=true Then
	If objFso.FileExists(strFilepath1& "open.ts")=true Then
		objFso.deleteFile strFilepath1 & "open.ts"
	End If
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & OpenClipName & " -c copy -bsf:v h264_mp4toannexb -f mpegts open.ts") '��Ƭͷmp4��ʽ����һ��.ts��ʽ
	WScript.echo oExec.StdErr.ReadAll()
End If
	
If objFso.FileExists(FinalClipName)=true  Then
	If objFso.FileExists(strFilepath1& "final.ts")=true Then
		objFso.deleteFile strFilepath1 & "final.ts"
	End If
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & FinalClipName & " -c copy -bsf:v h264_mp4toannexb -f mpegts final.ts") '��Ƭβmp4��ʽ����һ��.ts��ʽ
	WScript.echo oExec.StdErr.ReadAll()
End If

'������������Ƶ
If objFso.FileExists(strFilepath5)=true Then
	objFso.deleteFile strFilepath5
	WriteLine objPresentation.Name & "ɾ���ɹ�........."
End If
	
If objFso.FileExists(strFilepath1& "final.ts")=True Then
	If objFso.FileExists(strFilepath1& "open.ts")=True Then
		If objFso.FileExists(strFilepath1& "video.ts")=True Then
			Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i" & " " & chr(34) & "concat:open.ts|video.ts|final.ts" & chr(34) & " -c copy -bsf:a aac_adtstoasc -movflags +faststart" & " " & objPresentation.Name & ".mp4")
			'��Ƶ����Ƭͷ��Ƭβ
			WScript.echo oExec.StdErr.ReadAll()
			WriteLine objPresentation.Name & ".mp4" & "�ļ����ɳɹ�......."
		End If
	End If
End If
	
If objFso.FolderExists(strFilepath6)=false Then
	objFso.createfolder strFilepath6
	WriteLine "MP4�ļ��д����ɹ�"
End if

If objFso.FileExists(strFilepath5)=true Then
	If objFso.FileExists(strFilepath7)=true Then
		objFso.deleteFile strFilepath7
	End If
	objFso.MoveFile strFilepath5,strFilepath6 & "\\" '�����ɵ�������Ƶ�ƶ���MP4�ļ�����
End If

If objFso.FileExists( MP4File ) Then
	objFso.deleteFile MP4File
End If
	
objPresentation.Close
ObjPPT.Quit
WriteLine  "PPTת��Ƶ�ɹ�............"	

'�ر����г������̨
Dim WMI1,ProgList1,Prog1
Set WMI1=GetObject("winmgmts:\\.\root\cimv2")
Set ProgList1=WMI1.ExecQuery("Select * From Win32_Process")
For Each Prog1 In ProgList1
	If LCase(Prog1.Name)="cmd.exe" Then
		Prog1.Terminate
	Exit For
	End If
Next
