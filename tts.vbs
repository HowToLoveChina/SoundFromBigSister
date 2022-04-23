set fs =createobject("scripting.filesystemobject")
set ts=fs.opentextfile("tts.txt",1,true)
dim words(10000)
Randomize
idx = 1
do while ts.atendofstream <> true 
	x = ts.readline
	words(idx) = x 
	idx = idx + 1 
loop

do while true 
	str = ""
	for i = 1 to 3 
		x = int( idx * rnd + 1 )
		str = str + words(x)
	next 
	CreateObject("SAPI.SpVoice").Speak str 
loop 

