CREATE TABLE [dbo].[Patients]
(
	[Id] INT NOT NULL PRIMARY KEY IDENTITY, 
    [Number] INT NOT NULL, 
    [Surname] NVARCHAR(50) NOT NULL, 
    [Name] NVARCHAR(50) NOT NULL, 
    [Otchestvo] NVARCHAR(50) NOT NULL, 
    [Data_Rozhdeniya] DATE NOT NULL
)
