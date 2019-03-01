

; Sales Order
Text:="|<Sales Order>*135$73.zw0000000000E3000000000081E0000000004DY0040000102Tl1s2003k0U1CTV41002A0E0a0EUwbD22ttnns8Q1IIV1F558z41bfv0UcWyY7m0IJ0MEIFEG8N2+OUYAO9c9CQVtxDS3t3nobwE000000000Ew80000000008040000000007zy0000000002"
if ok:=FindText(122,57,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



; New Button on Sales Order Tab
Text:="|<New Button On Sales Order>*178$37.000000M00000A001000001E000Dtg0207V100k3yP00M1050000U000U0E080E08040BU40203k20100U100U000U0E000Q7s03000001U00000000000000000Dzzzzzw"
if ok:=FindText(238,130,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



; REFRESH on Sales Order
Text:="|<Refresh on Sales Order>*170$78.00000000000000000000000000U000008000000U00800M000001000I00s00000107wq01z007U0107V100zU08E01U7wq00Mk0EDU1U40I008M0E0U1040004080E0U1040004080Ezw1040206400F041U40203600G080U40201z00I0E0040200zU0M0U0040200700Dz00071y006000000U000004000000U00000000000000000000000000000000000000U"
if ok:=FindText(260,129,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}




; 자동 이메일 보내기 라디오 버튼 체크하기 위함
Text:="|<Auto Inv Email>*143$78.60100U007U00960300U0040001D4rb0bqM47zD9/4nAUaGE7aP99N4n8kaHk44F79TYn8kaHk44F99FanAUaFY44FN9krlb0aFY7oFDdU"
if ok:=FindText(936,336,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



; 자동 이메일 보내기 주소
Text:="|<Auto Inv Email Adr>*144$99.60100U00w00183080k0M0400400010M10D4rb0bqMUztt87VttMaNY4mG7aP990gPAN4n8kaHkUW8t8AW9XwaN64mS44F991yFAFanAUaFYUW/988nNa6yAs4mAboFDd33DAU"
if ok:=FindText(765,359,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



; Customer on Sales Order
Text:="|<Customer on Sales Order>*175$54.D00100000lU0100000UoFnb7r77U4GNAaNBaU4G18YF8YU4Fl8YFDYUoE98YF84lYm9AYFAYD7nlb4F74U"
if ok:=FindText(254,226,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}


; Customer on SO Manager
Text:="|<Customer on SO Manager>*131$45.D00E00068020000UFSxngQw2+GHGIIUFMG+GyY28mFGI4lHGHOGkXvvniGHoU"
if ok:=FindText(553,127,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}





; Customer on the top menu
Text:="|<Customer on the top menu>*144$86.000000000000000000000000000000000000000000007U0000000000Dk2600000000002010k0000000000U09k00000000008E3l1s020000002S0EFW00U000000bU00E8jStzCRs09U2242+GPGIKG02Q0C10WkYIZx600X7YEE8X959EEM081u26+OGPGK4G028VUkyywvYYx7U0bMTw00000000009rzz000000000027zzk0000000000Uzw00000000000Dk0000000000000000000000000000U"
if ok:=FindText(1969,56,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



















; Sales Order on the Menu bar
Text:="|<Sales Order on the Menu bar>*137$73.zw0000000000E3000000000081E0000000004DY0040000102Tl1s2007k0U1CTV41006A0E0a0EUwbD22ttnns8Q1IIV1F558z41bfv0UcWyY7m0IJ0MEIFEG8N2+OkYAO9g9CQVtxDS3t3nobwE000000000Ew80000000008040000000007zy0000000002"
if ok:=FindText(122,57,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}


; New Button
Text:="|<New Button of Sales Order>*174$35.000001U00003000E00001E000zak001sEE0k3yP01U40I000800080E080E0U0E0s100U0k201000402000804000Q7s0A00000M00000U"
if ok:=FindText(237,129,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}


; Add(+) Button
Text:="|<Add(+) Button of Sales Order>*147$14.0000030180G05UDz4ztTyDz0S07U1s0A000008"
if ok:=FindText(458,128,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}


; Save Button
Text:="|<Save Button of Sales Order>*167$58.k0k0w7U0TXU70401011j0w0G0Y041S7U182E0E4ww04U901TFzU0G0Y0w13w0182E2LoDk04zt090FzU0E040bzDD013sE201sS04EF080D0w0Fz40zzs1k14QE3zw0007lT0A00000Tzw0TzU"
if ok:=FindText(381,129,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}
















; Cumstomer Master 의 Edit에 있는 Auto Inv.Email Adr
Text:="|<Auto Inv.Email Adr>*144$99.60100U00w00183080k0M0400400010M10D4rb0bqMUztt87VttMaNY4mG7aP990gPAN4n8kaHkUW8t8AW9XwaN64mS44F991yFAFanAUaFYUW/988nNa6yAs4mAboFDd33DAU"
if ok:=FindText(765,359,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}



; Cumstomer Master 의 Edit에 있는 Auto Inv.Email 체크 하기 위함
Text:="|<Auto Inv.Email>*144$78.60100U007U00960300U0040001D4rb0bqM47zD9/4nAUaGE7aP99N4n8kaHk44F79TYn8kaHk44F99FanAUaFY44FN9krlb0aFY7oFDdU"
if ok:=FindText(936,336,150000,150000,0,0,Text)
{
  CoordMode, Mouse
  X:=ok.1, Y:=ok.2, W:=ok.3, H:=ok.4, Comment:=ok.5
  MouseMove, X+W//2, Y+H//2
}

