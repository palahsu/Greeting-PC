Set Sapi = Wscript.CreateObject("SAPI.SpVoice")

Function greeting()
	Sapi.speak "The current time is"
	
	if hour(time) > 12 then
 		Sapi.speak hour(time)-12
 	else 
 		if hour(time) = 0 then
 			Sapi.speak "12"
	 	else
	 		Sapi.speak hour(time)
	 	end if
 	end if

	if minute(time) < 10 then
 		Sapi.speak "o"	 	
	 	if minute(time) < 1 then
	 		Sapi.speak "clock"
	 	else
	 		Sapi.speak minute(time)
	 	end if 	
 	else
 		Sapi.speak minute(time)
 	end if
	
	if hour(time) > 12 then
 		Sapi.speak "P.M."
 	else
 		if hour(time) = 0 then
 		if minute(time) = 0 then
 			Sapi.speak "Midnight"
 		else
 			Sapi.speak "A.M."
 		end if
 		else
 			if hour(time) = 12 then
 			if minute(time) = 0 then
 			Sapi.speak "Noon"
 		else
			Sapi.speak "P.M."
 		  end if
 		else
 			Sapi.speak "A.M."
 		end if
 	end if
     end if
End Function

do

wscript.sleep 100

if hour(time) >= 10 AND hour(time) < 12 then
 Sapi.speak "Good Morning"
 Call greeting()
elseif hour(time) >= 12 AND hour(time) < 13 then
 Sapi.speak "Good Noon"
 Call greeting()
elseif hour(time) >= 14 AND hour(time) < 16 then
 Sapi.speak "Good After Noon"
 Call greeting()
elseif hour(time) >= 17 AND hour(time) < 22 then
 Sapi.speak "Good Evening"
 Call greeting()
elseif hour(time) >= 23 OR hour(time) >= 2 then
 sapi.speak " Good Night"
 Call greeting()
else
 sapi.speak " "
end if 

wscript.sleep 360000
loop