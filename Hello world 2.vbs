 Dim Message, Speak
      	Message=InputBox("Hello World","Speak")
      	Set Speak=CreateObject("sapi.spvoice")
      	Speak.Speak Message