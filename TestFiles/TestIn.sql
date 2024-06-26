-- == CsvInsertGen.py == -- 
-- Created       = 2022-12-02 15:52:20.331006
-- fileinname    = ./TestFiles/TestIn.xlsx
-- fileoutname   = ./TestFiles/TestIn.sql
-- tblname       = TestTable
-- Total rows in = 5 (+ 1 header)

SET NOCOUNT ON

IF EXISTS(Select * from sys.objects Where Name = 'TestTable' and Type = 'U')
BEGIN
    DROP TABLE TestTable
END
GO

CREATE TABLE TestTable
(
  _ID INT NOT NULL IDENTITY (1,1)
, [f1] NVARCHAR(50) NULL
, [f2] NVARCHAR(50) NULL
, [f3] NVARCHAR(50) NULL
, [f4] NVARCHAR(50) NULL
, [F5] NVARCHAR(250) NULL
)
GO

INSERT TestTable ([f1], [f2], [f3], [f4], [F5])
VALUES ('va1', 'va2', 'va3', 'Mary', 'Muffet')

INSERT TestTable ([f1], [f2], [f3], [f4], [F5])
VALUES ('vb1', 'vb2', 'vb3', 'Al', 'Santaballa')

INSERT TestTable ([f1], [f2], [f3], [f4], [F5])
VALUES ('vc1', 'vc2', 'vc3', 'John', 'Doe')

INSERT TestTable ([f1], [f2], [f3], [f4], [F5])
VALUES ('vd1', 'vd2', 'vd3', 'Joan', 'VeryVeryLasongLastNameEvenLongerThanRumplestilskinOrMxysptlkOrEvenSantaballaVeryVeryLasongLastNameEvenLongerThanRumplestilskinOrMxysptlkOrEvenSantaballaVeryVeryLasongLastNameEvenLongerThanRumplestilskinOrMxysptlkOrEvenSantaballa')

SELECT Cnt = COUNT(*) FROM TestTable
GO

-- Total to insert 4
