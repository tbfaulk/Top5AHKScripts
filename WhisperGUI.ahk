;Authored by Tyler Faulkner
;Special thanks to the fine folks at The Automator and the AHK Hero Group

;Note that in order to use this script, you must have already installed Whisper
;This link is a great resource if you have not yet installed Whisper:  https://youtu.be/ABFqbY_rmEk

#Requires Autohotkey v2.0+
#SingleInstance Force

Main := Gui(,"Whisper")

;File Selection
Main.Add("Button", "xm","Select File" ).OnEvent("Click", FileSelect_Click) 
Main.Add("Text", "w300 x+10 vGuiFiles")

;Text options
Main.Add("Text", "xm" ,"Output Format")
OutputFormat := Main.Add("DropDownList", "w60 x100 y+-15 Choose2 vOutputFormat" , ["ALL","txt","json","srt","tsv","vtt"])

;Model options
Main.Add("Text", "xm y+5","Model Type")
Model := Main.Add("DropDownList", "w60 x100 y+-15 Choose4 vModel" , ["Large", "Medium", "Small", "Base", "Tiny"])

;Buttons
Main.Add("Button", "default w75", "OK").OnEvent("Click",Ok_Click) 
Main.Add("Button", "w75 x+m" , "Cancel").OnEvent("Click",Cancel_Click) 


FileSelect_Click(GuiCtrl, *)
{
    SelectedFile := FileSelect(, , "Select File to Transcribe", "Audio (*.mp3; *.mp4; *.m4a; *.mpeg; *.mpga; *.wav; *.webm)" )
    ; Access the GUI via a reference to the button control and then the target control by its variable / name / handle)
    GuiCtrl.Gui["GuiFiles"].Text := SelectedFile
}

Cancel_Click(*)
{
Main.Destroy()
}

Ok_Click(*)
{
;msgbox Main["GuiFiles"].value "`n" Main["OutputFormat"].text "`n" Main["Model"].text
if Main["GuiFiles"].value
    {
    Splitpath Main["GuiFiles"].value, &GuiFilesname, &GuiFilesdir
    Prompt := "/c whisper" ' "' GuiFilesname '"' " --model " StrLower(Main["Model"].text) ".en --output_format " StrLower(Main["OutputFormat"].text)
    RunWait "cmd.exe '" Prompt, GuiFilesdir, ""
    Run GuiFilesdir
    Main.Destroy()
    } else{
        Msgbox "Select a file to continue."    
    }
}

Main.Show("w400")
