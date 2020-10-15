// Create a variable speaks which will be used to say things through microsoft voice or voice command module
Dim speaks, speaks2, speaks3, speech, input_command

// speaks variable which will say or voice command which will be created when the windows starts
speaks="Welcome sir! Jarvis at your service can i do for you?  "// added spaces to the last to have a pause in between the voice operations speaks2
speaks2=" Please sir say the command and I will do as you wish  " // added spaces at last to have a pause


// Create object speech and call methond Speak
Set speech=CreateObject("sapi.spvoice")
speech.Speak speaks

//capture  the input from microphone of the user and do the work or obey the command
input_command = MicrophoneCaptureVoice()
  //function to obey the command or do the action in windows
WindowsCommandineInterface.do("input_command")
  
end
