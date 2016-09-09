// Create a variable speaks 
Dim speaks, speech

// Change the value of the variable which you want your pc to say
speaks="Welcome sir! What can i do for you?"

// Create object speech and call methond Speak
Set speech=CreateObject("sapi.spvoice")
speech.Speak speaks
