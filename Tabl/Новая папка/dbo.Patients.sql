CREATE TABLE [dbo].[Patients] (
    [Id]              INT           IDENTITY (1, 1) NOT NULL,
    [Num]             NVARCHAR (50) NOT NULL,
    [Surname]         NVARCHAR (50) NOT NULL,
    [Name]            NVARCHAR (50) NOT NULL,
    [Otchestvo]       NVARCHAR (50) NOT NULL,
    [Data_Rozhdeniya] DATE          NULL,
    [SurnameD]        NVARCHAR (50) NOT NULL,
    [NameD]           NVARCHAR (50) NOT NULL,
    [OtchestvoD]      NVARCHAR (50) NOT NULL,
    [Dolzhnost]       NVARCHAR (50) NOT NULL,
    [Ill]             NVARCHAR (50) NOT NULL,
    [Heal]            NVARCHAR (50) NOT NULL,
    [Time]            NVARCHAR (50) NOT NULL,
    [Dispanser]       NVARCHAR (50) NOT NULL,
    [Info]            NVARCHAR (50) NOT NULL,
    PRIMARY KEY CLUSTERED ([Id] ASC)
);

