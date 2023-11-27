Attribute VB_Name = "DF_Gun"
Option Explicit

Sub InitAAGuns()
'AA Gun Coordinates AA_SQ/JA(X, Y) from upper left corner of board clockwise from SQA2 outer gun
AA_SQ(1, 1) = 13: AA_SQ(1, 2) = 9
AA_SQ(2, 1) = 12: AA_SQ(2, 2) = 9
AA_SQ(3, 1) = 11: AA_SQ(3, 2) = 10
AA_SQ(4, 1) = 11: AA_SQ(4, 2) = 11

AA_SQ(5, 1) = 11: AA_SQ(5, 2) = 2
AA_SQ(6, 1) = 11: AA_SQ(6, 2) = 3
AA_SQ(7, 1) = 12: AA_SQ(7, 2) = 4
AA_SQ(8, 1) = 13: AA_SQ(8, 2) = 4

AA_JA(1, 1) = 4: AA_JA(1, 2) = 2
AA_JA(2, 1) = 4: AA_JA(2, 2) = 3
AA_JA(3, 1) = 3: AA_JA(3, 2) = 4
AA_JA(4, 1) = 2: AA_JA(4, 2) = 4

AA_JA(5, 1) = 2: AA_JA(5, 2) = 9
AA_JA(6, 1) = 3: AA_JA(6, 2) = 9
AA_JA(7, 1) = 4: AA_JA(7, 2) = 10
AA_JA(8, 1) = 4: AA_JA(8, 2) = 11

End Sub
