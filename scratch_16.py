="The stock is currently trading at " & TEXT(C2,"$####.00") & ", " & IF(C2-C7=0, "at the same price" & " after opening "
& IF(C8-C7=0, "at the same price as yesterday's close", IF(C8-C7>0, "up " & IF((C7-C8)/C7*-1 <0.01, "slightly", TEXT((C7-C8)/C7*-1,"##.##%")) & " over yesterday's close",
IF((C7-C8)/C7 <0.01, "slightly below", "down " & TEXT((C7-C8)/C7*1,"##.##%") & ) & " yesterday's close")), IF(C2-C7>0, "up " & TEXT((C7-C2)/C7*-1,"##.##%") & " after opening "
& IF(C8-C7=0, "at the same price as yesterday's close", IF(C8-C7>0, "up " & IF((C7-C8)/C7*-1 <0.01, "slightly", TEXT((C7-C8)/C7*-1,"##.##%")) & " over yesterday's close",
IF((C7-C8)/C7 <0.01, "slightly below", "down from" & TEXT((C7-C8)/C7*1,"##.##%")) & " yesterday's close")), "down " & TEXT((C7-C2)/C7*1,"##.##%") & " after opening " &
IF(C8-C7=0, "at the same price as yesterday's close", IF(C8-C7>0, "up " & IF((C7-C8)/C7*-1 <0.01, "slightly", TEXT((C7-C8)/C7*-1,"##.##%")) & " over yesterday's close",
IF((C7-C8)/C7 <0.01, "slightly below", "down from" & TEXT((C7-C8)/C7*1,"##.##%")) & " yesterday's close")) ))