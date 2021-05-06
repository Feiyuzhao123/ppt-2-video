' Write each slide's notes to a text file
' in same directory as presentation itself
' Each file is named NNNN_Notes_Slide_xxx
' where NNNN is the name of the presentation
'       xxx is the slide number
'PPT播放的时间不要超过二十分钟
'PPT不能设置切换效果，比如浅入....
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
ppPlaceholderBody = 2  '代表 Body
Dim dub_volume         '配音音量
Dim background_volume  '背景音乐音量
Dim CurrentPath
dub_volume = 1
background_volume = -53

Sub WriteLine ( strLine )
    WScript.Stdout.WriteLine strLine
End Sub

Set WshShell = WScript.CreateObject("WScript.Shell") 

Path = WshShell.CurrentDirectory    '获取脚本路径
WriteLine "脚本的路径为 " & Path 

CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
WriteLine "当前工作目录是：" & CurrentPath

If WScript.Arguments.Count < 1 Then
	WriteLine "PPT to Video v2.0"
	WriteLine "cscript ppt2video PPT/PPTx [t/T]"
	WriteLine "使用说明：You need to specify input PPT/PPTx file。[t/T]使用t或者T，表示仅仅是插入合成音频，不进行后续的转成视频处理。"
	WScript.Quit
End If

'############################################################################################################################################
inputFile = CurrentPath & "\\pptx\\" & WScript.Arguments(0)	'要打开PPT文件的地址
strFilepath6  = CurrentPath & "\\" & "MP4"				'存放MP4的文件夹路径
'############################################################################################################################################

Set objFso = CreateObject("Scripting.FileSystemObject")

If objFso.FileExists( inputFile ) = False Then   '查找PPT是否存在
	WriteLine "找不到PPT/PPTx文件：" & inputFile
	WScript.Quit
End If

Set objPPT = CreateObject( "PowerPoint.Application" )

objPPT.Visible = True

Set objPres = objPPT.Presentations.Open(inputFile)

Set objPresentation = objPPT.ActivePresentation 


Dim strFilepath1,strFilepath2,strFilepath3,strFilepath4,strFilepath5

Dim strSrcFld, strDstFile

strFilepath1 = Path & "\\"  
strFilepath7  = strFilepath6 & "\\"& objPresentation.Name & ".mp4"      '最后存放MP4的文件路径                                 

'这里面的地址不能随意改动，他需要在ffmpeg命令下执行，需要和ffmpeg安装包放在同一个文件夹下

OpenClipName  = Path & "\\" & "open.mp4"                      '片头路径
FinalClipName = Path & "\\" & "final.mp4"                     '片尾路径 
strFilepath2 =  Path & "\\" & "wav" & "\\"                    'PPT所在的文件路径下wav文件夹下的文件（e:\PPT2Video\wav\）
strFilepath3 = strFilepath2 & "*.wav"                         'wav文件夹下所有的.wav文件（e:\PPT2Video\wav\*.wav）
strFilepath4  = strFilepath1 & "back" & "\\" & "backg.wav"    '背景音乐的路径
strFilepath5  = Path  & "\\" & objPresentation.Name & ".mp4"
strSrcFld = strFilepath2                                      '要提取.wav文件名的路径
strDstFile = strFilepath1 & "list.txt"                        '提取文件名生成list.txt的路径

' One --------------------------------------------------------------- 第一步：提取PPT中文字，生成文件

WriteLine "****************提取PPT中文字****************"

'删除strFilepath1中的所有.txt和.wav文件，strFilepath2中的所有.wav文件

	if objFso.fileexists(strFilepath1 & "1.txt")=True then
		objFso.deleteFile strFilepath1 & "???.txt"
		WriteLine "所有的txt文字文件删除成功........."
	End If
	
	if objFso.fileexists(strFilepath1 & "001.wav")=True then
		objFso.deleteFile strFilepath1 & "*.wav"
	End If
	
	if objFso.fileexists(strFilepath2 & "001.wav")=True then 
		objFso.deleteFile strFilepath2 & "*.wav"
		WriteLine "所有的wav音频文件删除成功........"
	End If

'提取文字，生成.txt文件

Dim SlideNumber
'Dim cc
SlideNumber = 1

	For Each oSlide In objPresentation.Slides
		WriteLine("获取PPT的页数：" & "第" & SlideNumber & "页" & "........")	
	
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

' Two ---------------------------------------------------------------第二步：文字转语音

WriteLine "****************文字转语音****************"

If objFso.FolderExists(Path & "\\" & "wav")=false Then
	objFso.createfolder Path & "\\" & "wav"
	WriteLine "wav文件夹创建成功........"
End if

Dim objShell

Dim objWMIService, colProcessList, strComputer

strComputer = "."

Set objShell = CreateObject("Wscript.Shell")

objShell.Run "tts.exe" '调用tts.exe，进行语音转换

Do While True
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'tts.exe'")
	IF colProcessList.Count > 0 THEN
	Else
		Exit Do	
	END IF
Loop

' Three ---------------------------------------------------------------第三步：改变PPT切换动画时间

WriteLine "****************改变PPT切换动画时间****************"

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
	
'读取页数，页数不大于9时，设置为001,002.......，当页数大于9，则设置为010,011........。
'为了方便读取.wav文件名时，按照顺序排列，方便语音连接，生成配音。	

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
		'获取音频时间
		WScript.echo oExec.StdErr.ReadAll()
		Return_Value = oExec.StdOut.ReadLine
		' add by lujun 如果讯飞返回错误，也就是TTS生成的wav文件无法返回有效长度，那么缺省把每页的切换时间设置为 1，这样避免程序错误退出。
		If Return_Value = "N/A" Then
			Return_Value = 1
		End If
		WriteLine "获取PPT切换的时间(Slide Num " & SlideNumber & "): " & Return_Value	& "秒........"
		oSlide.SlideShowTransition.AdvanceOnClick = False
		oSlide.SlideShowTransition.AdvanceOnTime = True
		'oSlide.SlideShowTransition.Duration = 2
		'---------------------------------------------------------------------------------------	
		oSlide.SlideShowTransition.AdvanceTime = Return_Value			' 修改每页的持续播放时间
		'---------------------------------------------------------------------------------------			
		
		'删除原来增加的声音对象
		Dim intShape
		With oSlide.Shapes
			For intShape = .Count To 1 Step -1
				With .Item(intShape)
					If .Type = 16 And .Name = "PPT2Video" Then .Delete	' 16 代表使用 Shapes.AddMediaObject2 增加的声音对象，Name判断是为了防止误删除其它的声音对象
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
					.StopAfterSlides = 9999	'让声音可以与动画一起播放
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
		WriteLine  "PPT中添加合成的发音音频成功............"
		WScript.Quit
	End if
End If
'four---------------------------------------------------------------第四步：将PPT转为MP4

WriteLine "****************PPT转为视频****************"

'objPresentation.CreateVideo(fileName, DefaultSlideDuration:=1, VertResolution:=480)
'可以工作的pdf objPresentation.SaveAs "c:\\lujun.mp4", 32, False

'注意！由于生成MP4文件要耗费相当多的时间，因此后面的 objPresentation.Close 与 ObjPPT.Quit不能马上执行，
'否则会导致生成MP4文件出错。正确的做法是等待MP4生成完毕后，才执行 objPresentation.Close 与 ObjPPT.Quit。

Dim ppSaveAsMP4,f,MP4File
ppSaveAsMP4 = 39   '39代表以mp4格式输出的参数
MP4File = Path & "\" & objPresentation.Name & ".save.mp4"	'生成的mp4路径

If objFso.FileExists( MP4File ) Then
	objFso.deleteFile MP4File
End If

If WScript.Arguments.Count = 2 Then
	'CreateVideo( FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertResolution, FramesPerSecond, Quality )
	'UseTimingsAndNarrations:=True, VertResolution:=720, FramesPerSecond:=30, Quality:=100
	'WScript.Arguments(1)是120、240、480等
	objPresentation.CreateVideo MP4File,True,3,WScript.Arguments(1),30,60
	'ppMediaTaskStatusNone				0
	'ppMediaTaskStatusInProgress		1
	'ppMediaTaskStatusQueued			2
	'ppMediaTaskStatusDone				3
	'ppMediaTaskStatusFailed			4
	'等待PPT转换为MP4完成
	Do While True
		wscript.sleep 1000
		If objPresentation.CreateVideoStatus = 3 Or objPresentation.CreateVideoStatus = 4 Then
			Exit Do
		End if
	Loop
Else
	'------下面这行转换代码，转换的视频质量高，但是耗时很长
	objPresentation.SaveAs MP4File , ppSaveAsMP4, False			'将ppt转为视频，并以mp4格式保存
End If

'等待PPT转换为MP4完成
Do While True	
	If objFso.FileExists( MP4File ) = False Then			'这里必须要wscript.sleep，不然找不到视频的路径
		wscript.sleep 3000
	Else
		wscript.sleep 3000
		Set f = objFso.GetFile(MP4File)  '获取视频的大小
		If f.Size > 0 Then
			WriteLine MP4File & " 文件大小为 :" & f.Size & " Bytes........"
			Exit Do
		End if
	End If
Loop

Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & MP4File & " -vn -y -acodec copy video.aac")'将音频从视频中分离出来
WScript.echo oExec.StdErr.ReadAll()
If objFso.FileExists( strFilepath1 & "video.aac" ) Then
	WriteLine "音频从视频中分离成功"
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i video.aac video.wav") '将音频格式转为.wav
	WScript.echo oExec.StdErr.ReadAll()
	objFso.deleteFile strFilepath1 & "video.aac"
	If objFso.FileExists( strFilepath1 & "video.aac" ) = False Then
		WriteLine "video.aac文件删除成功"
	Else
		WriteLine "video.aac文件删除失败"
	end if
Else
	WriteLine "音频从视频中分离失败，video.aac文件不存在"
End If

If objFso.FileExists( strFilepath1 & "video.wav" ) = False Then
	WriteLine "video.wav文件不存在"
Else
	WriteLine "video.wav文件存在"
End If
'five...................................................................第五步：将生成的所有语音连接生成配音

WriteLine "****************语音连接生成配音****************"

'提取wav文件夹中.wav文件的文件名，以file '1.wav'格式保存在list.txt文件中

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
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -f concat -i list.txt -c copy notes.wav") '连接所有的.wav文件，生成视频配音
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "配音生成成功........"
Else
	WriteLine "配音不存在........"
End If 
	
'six...................................................................第六步：将配音与背景音乐合并	

WriteLine "****************配音与背景音乐合并****************"

If objFso.FileExists(strFilepath1 & "dnotes.wav")=false Then
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i notes.wav -af volume=" & dub_volume & "dB dnotes.wav")  '增加配音音量
	WScript.echo oExec.StdErr.ReadAll()
End If
	
If objFso.FileExists(strFilepath1& "backg.wav")=false Then
	objFso.CopyFile strFilepath4,strFilepath1,False
End If

Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i backg.wav -af volume=" & background_volume & "dB dbackg.wav")  '增加背景音乐音量

WScript.echo oExec.StdErr.ReadAll()

if objFso.fileexists(strFilepath1 & "allvoice.mp3")=False then 
	WriteLine "配音和背景音乐合成文件不存在........"
else
	objFso.deleteFile strFilepath1 & "*.mp3"
	WriteLine "配音和背景音乐合成文件删除成功........"
End If
Dim v
If objFso.FileExists( strFilepath1 & "video.wav" ) = true Then
	WriteLine "video.wav文件存在........"
	Set v = objFso.GetFile(strFilepath1 & "video.wav")
	WriteLine "分离的音频文件大小为 :" & v.Size & " Bytes........"
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i dnotes.wav -i dbackg.wav -i video.wav -filter_complex amix=inputs=3:duration=first:dropout_transition=2 -f mp3 allvoice.mp3")
	'背景音乐与配音的合并
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "分离的音频与背景和配音合成成功"
else
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i dnotes.wav -i dbackg.wav -filter_complex amix=inputs=2:duration=first:dropout_transition=2 -f mp3 allvoice.mp3")
	'背景音乐与配音的合并
	WScript.echo oExec.StdErr.ReadAll()
	WriteLine "video.wav文件不存在........"
	WriteLine "配音和背景音乐合成成功........"
End if
'seven...................................................................第七步：将背景音乐与配音生成的音频与视频合并

WriteLine "****************音频与视频合并****************"

If objFso.FileExists(strFilepath1 & "video.mp4") Then
	objFso.deleteFile strFilepath1 & "video.mp4"
	WriteLine "video.mp4删除成功"
else
	WriteLine "video.mp4文件不存在"
End If
If objFso.FileExists( strFilepath1 & "video.wav" ) = true Then
	WriteLine "video.wav文件存在"
	If objFso.FileExists(strFilepath1 & "partvideo.mp4")=true Then
		objFso.deleteFile strFilepath1 & "partvideo.mp4"
	else
		WriteLine "partvideo.mp4文件不存在"
	end if
	If objFso.FileExists(strFilepath1 & "partvideo.mp4")=false Then
		If objFso.FileExists(MP4File) Then
			WriteLine MP4File & "存在"
			Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i " & MP4File & " -an partvideo.mp4") '从视频中提取视频，不要音频
			WScript.echo oExec.StdErr.ReadAll()
			If objFso.FileExists( strFilepath1 & "partvideo.mp4" )Then
				WriteLine "partvideo.mp4生成成功"
				If objFso.FileExists( strFilepath1 & "video.mp4" ) = False Then
					Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i partvideo.mp4 -i allvoice.mp3 -r 25 video.mp4") '音频与视频的合并
					WScript.echo oExec.StdErr.ReadAll()
					objFso.deleteFile strFilepath1 & "partvideo.mp4"
					if objFso.FileExists(strFilepath1 & "partvideo.mp4")=false Then
						WriteLine "partvideo.mp4删除成功"
					else
						WriteLine "partvideo.mp4删除失败"
					end if
				else 
					WriteLine "video.mp4文件存在"
				end if 
			else
				WriteLine "partvideo.mp4生成失败"
			end if
		else
			WriteLine MP4File & "不存在"
		end if
	else
		WriteLine "partvideo.mp4文件存在"
	end if
else
	if objFso.FileExists(strFilepath1 & "video.mp4")=false Then
		Set oExec = WshShell.Exec("cmd /c " & strFilepath1 & "ffmpeg -i " & MP4File & " -i allvoice.mp3 -r 25 video.mp4") '音频与视频的合并
		WScript.echo oExec.StdErr.ReadAll()
		WriteLine "背景音乐，配音与视频合成成功........"
	Else 
		WriteLine "video.mp4文件不存在"
	End if 
End if


'eight.....................................................................第八步:片头，片尾与视频合并	

WriteLine "*****************片头，片尾与视频合并****************"

If objFso.FileExists(strFilepath1 & "video.ts")=true Then
	objFso.deleteFile strFilepath1 & "video.ts"
End If
	
Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i video.mp4 -c copy -bsf:v h264_mp4toannexb -f mpegts video.ts") '将视频生成一个.ts格式，方便视频的连接

WScript.echo oExec.StdErr.ReadAll()

If objFso.FileExists(OpenClipName)=true Then
	If objFso.FileExists(strFilepath1& "open.ts")=true Then
		objFso.deleteFile strFilepath1 & "open.ts"
	End If
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & OpenClipName & " -c copy -bsf:v h264_mp4toannexb -f mpegts open.ts") '将片头mp4格式生成一个.ts格式
	WScript.echo oExec.StdErr.ReadAll()
End If
	
If objFso.FileExists(FinalClipName)=true  Then
	If objFso.FileExists(strFilepath1& "final.ts")=true Then
		objFso.deleteFile strFilepath1 & "final.ts"
	End If
	Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i " & FinalClipName & " -c copy -bsf:v h264_mp4toannexb -f mpegts final.ts") '将片尾mp4格式生成一个.ts格式
	WScript.echo oExec.StdErr.ReadAll()
End If

'制作完整的视频
If objFso.FileExists(strFilepath5)=true Then
	objFso.deleteFile strFilepath5
	WriteLine objPresentation.Name & "删除成功........."
End If
	
If objFso.FileExists(strFilepath1& "final.ts")=True Then
	If objFso.FileExists(strFilepath1& "open.ts")=True Then
		If objFso.FileExists(strFilepath1& "video.ts")=True Then
			Set oExec = WshShell.Exec("cmd /c" & strFilepath1 & "ffmpeg -i" & " " & chr(34) & "concat:open.ts|video.ts|final.ts" & chr(34) & " -c copy -bsf:a aac_adtstoasc -movflags +faststart" & " " & objPresentation.Name & ".mp4")
			'视频连接片头与片尾
			WScript.echo oExec.StdErr.ReadAll()
			WriteLine objPresentation.Name & ".mp4" & "文件生成成功......."
		End If
	End If
End If
	
If objFso.FolderExists(strFilepath6)=false Then
	objFso.createfolder strFilepath6
	WriteLine "MP4文件夹创建成功"
End if

If objFso.FileExists(strFilepath5)=true Then
	If objFso.FileExists(strFilepath7)=true Then
		objFso.deleteFile strFilepath7
	End If
	objFso.MoveFile strFilepath5,strFilepath6 & "\\" '将生成的完整视频移动到MP4文件夹中
End If

If objFso.FileExists( MP4File ) Then
	objFso.deleteFile MP4File
End If
	
objPresentation.Close
ObjPPT.Quit
WriteLine  "PPT转视频成功............"	

'关闭运行程序控制台
Dim WMI1,ProgList1,Prog1
Set WMI1=GetObject("winmgmts:\\.\root\cimv2")
Set ProgList1=WMI1.ExecQuery("Select * From Win32_Process")
For Each Prog1 In ProgList1
	If LCase(Prog1.Name)="cmd.exe" Then
		Prog1.Terminate
	Exit For
	End If
Next
