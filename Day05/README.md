not sure there was a good way to do today's in Excel, fair warning may cause excel to pull all resources (specifically the p1/p2 files)

A2 = X1, B2 = Y1, C2 = X2, D2 = Y2, H1 an index of line length (up to 990 being the highest coordinate)

=IF($F2,IF(AND($A2=$C2,H$1<=ABS($B2-$D2),$B2>$D2),$A2&","&$B2-H$1,IF(AND($A2=$C2,H$1<=ABS($B2-$D2),$D2>$B2),$A2&","&$D2-H$1,IF(AND($B2=$D2,H$1<=ABS($A2-$C2),$A2>$C2),$A2-H$1&","&$B2,IF(AND($B2=$D2,H$1<=ABS($A2-$C2),$C2>$A2),$C2-H$1&","&$B2,"")))),IF(H$1<=ABS($B2-$D2),IF(AND($A2>$C2,$B2>$D2),$A2-H$1&","&$B2-H$1,IF(AND($C2>$A2,$D2>$B2),$C2-H$1&","&$D2-H$1,IF(AND($A2>$C2,$B2<$D2),$A2-H$1&","&$B2+H$1,$A2+H$1&","&$B2-H$1))),""))
