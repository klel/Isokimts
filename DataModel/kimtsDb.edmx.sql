
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 02/03/2015 15:04:32
-- Generated from EDMX file: C:\Users\ken\Dropbox\kimts\kimts\DataModel\kimtsDb.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [okimtsDb];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_InboxDocuments_BuildinObjects]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[InboxDocuments] DROP CONSTRAINT [FK_InboxDocuments_BuildinObjects];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_BuildinObjects]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_BuildinObjects];
GO
IF OBJECT_ID(N'[dbo].[FK_ContractorEmployes_Contractors]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[ContractorEmployes] DROP CONSTRAINT [FK_ContractorEmployes_Contractors];
GO
IF OBJECT_ID(N'[dbo].[FK_InboxDocuments_ContractorEmployes]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[InboxDocuments] DROP CONSTRAINT [FK_InboxDocuments_ContractorEmployes];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_ContractorEmployes]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_ContractorEmployes];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_Contractors]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_Contractors];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_DocStates]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_DocStates];
GO
IF OBJECT_ID(N'[dbo].[FK_InboxDocuments_Employes]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[InboxDocuments] DROP CONSTRAINT [FK_InboxDocuments_Employes];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_EmployesWhoMade]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_EmployesWhoMade];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_EmployesWhoSign]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_EmployesWhoSign];
GO
IF OBJECT_ID(N'[dbo].[FK_Users_Employes]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Users] DROP CONSTRAINT [FK_Users_Employes];
GO
IF OBJECT_ID(N'[dbo].[FK_InboxDocuments_FileMetaData]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[InboxDocuments] DROP CONSTRAINT [FK_InboxDocuments_FileMetaData];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_FileMetaData]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_FileMetaData];
GO
IF OBJECT_ID(N'[dbo].[FK_InboxDocs_OutboxDocument]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[InboxDocuments] DROP CONSTRAINT [FK_InboxDocs_OutboxDocument];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_InboxDocs]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_InboxDocs];
GO
IF OBJECT_ID(N'[dbo].[FK_OutboxDocument_TypesOfOutboxDocs]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[OutboxDocuments] DROP CONSTRAINT [FK_OutboxDocument_TypesOfOutboxDocs];
GO
IF OBJECT_ID(N'[dbo].[FK_CompanyStructure_CompanyStructure]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[CompanyStructures] DROP CONSTRAINT [FK_CompanyStructure_CompanyStructure];
GO
IF OBJECT_ID(N'[dbo].[FK_Employes_CompanyStructure]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Employes] DROP CONSTRAINT [FK_Employes_CompanyStructure];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[BuildinObjects]', 'U') IS NOT NULL
    DROP TABLE [dbo].[BuildinObjects];
GO
IF OBJECT_ID(N'[dbo].[ContractorEmployes]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ContractorEmployes];
GO
IF OBJECT_ID(N'[dbo].[Contractors]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Contractors];
GO
IF OBJECT_ID(N'[dbo].[DocStates]', 'U') IS NOT NULL
    DROP TABLE [dbo].[DocStates];
GO
IF OBJECT_ID(N'[dbo].[Employes]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Employes];
GO
IF OBJECT_ID(N'[dbo].[FileMetaDatas]', 'U') IS NOT NULL
    DROP TABLE [dbo].[FileMetaDatas];
GO
IF OBJECT_ID(N'[dbo].[InboxDocuments]', 'U') IS NOT NULL
    DROP TABLE [dbo].[InboxDocuments];
GO
IF OBJECT_ID(N'[dbo].[OutboxDocuments]', 'U') IS NOT NULL
    DROP TABLE [dbo].[OutboxDocuments];
GO
IF OBJECT_ID(N'[dbo].[TypesOfOutboxDocs]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TypesOfOutboxDocs];
GO
IF OBJECT_ID(N'[dbo].[Users]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Users];
GO
IF OBJECT_ID(N'[dbo].[CompanyStructures]', 'U') IS NOT NULL
    DROP TABLE [dbo].[CompanyStructures];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'BuildinObjects'
CREATE TABLE [dbo].[BuildinObjects] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [ObjName] nchar(200)  NOT NULL,
    [ObjCode] nchar(10)  NULL,
    [ProjCypher] nchar(10)  NULL,
    [SZS] nchar(10)  NULL
);
GO

-- Creating table 'ContractorEmployes'
CREATE TABLE [dbo].[ContractorEmployes] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [FullName] nchar(400)  NOT NULL,
    [FName] nchar(100)  NOT NULL,
    [SName] nchar(100)  NOT NULL,
    [LName] nchar(100)  NOT NULL,
    [LNameDateln] nchar(100)  NULL,
    [Post] nchar(200)  NOT NULL,
    [PostDateln] nchar(200)  NULL,
    [Gender] nchar(10)  NOT NULL,
    [Email] nchar(100)  NULL,
    [Phone] nchar(50)  NULL,
    [Contractor] int  NULL
);
GO

-- Creating table 'Contractors'
CREATE TABLE [dbo].[Contractors] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [OrgName] nchar(500)  NOT NULL,
    [Email] nchar(100)  NULL,
    [Phone] nchar(25)  NULL,
    [Fax] nchar(25)  NULL
);
GO

-- Creating table 'DocStates'
CREATE TABLE [dbo].[DocStates] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [StateName] nchar(100)  NULL
);
GO

-- Creating table 'Employes'
CREATE TABLE [dbo].[Employes] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [FullName] nchar(200)  NOT NULL,
    [NameInDocs] nchar(150)  NULL,
    [FName] nchar(50)  NOT NULL,
    [SName] nchar(80)  NOT NULL,
    [LName] nchar(70)  NOT NULL,
    [State] nchar(10)  NULL,
    [Email] nchar(100)  NULL,
    [Phone] nchar(30)  NULL,
    [Post] nchar(200)  NULL,
    [DepartmentId] int  NOT NULL
);
GO

-- Creating table 'FileMetaDatas'
CREATE TABLE [dbo].[FileMetaDatas] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [FileName] nchar(80)  NOT NULL,
    [UploadDataTime] datetime  NOT NULL,
    [Users] int  NOT NULL,
    [FileSize] nchar(50)  NULL,
    [Data] varbinary(max)  NULL
);
GO

-- Creating table 'InboxDocuments'
CREATE TABLE [dbo].[InboxDocuments] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Sender] int  NOT NULL,
    [SenderNum] nchar(20)  NOT NULL,
    [SenderDate] datetime  NOT NULL,
    [ResponseOn] int  NULL,
    [Reciever] int  NOT NULL,
    [DocTheme] nchar(100)  NOT NULL,
    [DocState] int  NULL,
    [BuildingObj] int  NULL,
    [Files] int  NULL
);
GO

-- Creating table 'OutboxDocuments'
CREATE TABLE [dbo].[OutboxDocuments] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [RecieverOrg] int  NOT NULL,
    [RecieverEmploye] int  NULL,
    [BuildingObj] int  NULL,
    [OutboxNum] nchar(10)  NULL,
    [OutboxDate] datetime  NULL,
    [DocTheme] nchar(60)  NOT NULL,
    [WhoSign] int  NOT NULL,
    [WhoMade] int  NOT NULL,
    [ResponseOn] int  NULL,
    [SentDate] datetime  NULL,
    [DocState] int  NULL,
    [Files] int  NULL,
    [TypeOfOutboxDoc] int  NOT NULL,
    [IndexNumber] int  NULL
);
GO

-- Creating table 'TypesOfOutboxDocs'
CREATE TABLE [dbo].[TypesOfOutboxDocs] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [NameOfType] nchar(180)  NOT NULL,
    [NumberingMethod] nchar(40)  NOT NULL,
    [NumeratorPrefix] nchar(10)  NULL
);
GO

-- Creating table 'Users'
CREATE TABLE [dbo].[Users] (
    [id] int IDENTITY(1,1) NOT NULL,
    [UniqId] uniqueidentifier  NOT NULL,
    [Name] nchar(100)  NULL,
    [Employe] int  NOT NULL,
    [Invalid] bit  NULL,
    [Files] int  NULL
);
GO

-- Creating table 'CompanyStructures'
CREATE TABLE [dbo].[CompanyStructures] (
    [Id] int  NOT NULL,
    [DepartmentName] nchar(200)  NOT NULL,
    [ParentDepartmentId] int  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'BuildinObjects'
ALTER TABLE [dbo].[BuildinObjects]
ADD CONSTRAINT [PK_BuildinObjects]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'ContractorEmployes'
ALTER TABLE [dbo].[ContractorEmployes]
ADD CONSTRAINT [PK_ContractorEmployes]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Contractors'
ALTER TABLE [dbo].[Contractors]
ADD CONSTRAINT [PK_Contractors]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'DocStates'
ALTER TABLE [dbo].[DocStates]
ADD CONSTRAINT [PK_DocStates]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'Employes'
ALTER TABLE [dbo].[Employes]
ADD CONSTRAINT [PK_Employes]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'FileMetaDatas'
ALTER TABLE [dbo].[FileMetaDatas]
ADD CONSTRAINT [PK_FileMetaDatas]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [PK_InboxDocuments]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [PK_OutboxDocuments]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TypesOfOutboxDocs'
ALTER TABLE [dbo].[TypesOfOutboxDocs]
ADD CONSTRAINT [PK_TypesOfOutboxDocs]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [id] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [PK_Users]
    PRIMARY KEY CLUSTERED ([id] ASC);
GO

-- Creating primary key on [Id] in table 'CompanyStructures'
ALTER TABLE [dbo].[CompanyStructures]
ADD CONSTRAINT [PK_CompanyStructures]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [BuildingObj] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [FK_InboxDocuments_BuildinObjects]
    FOREIGN KEY ([BuildingObj])
    REFERENCES [dbo].[BuildinObjects]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_InboxDocuments_BuildinObjects'
CREATE INDEX [IX_FK_InboxDocuments_BuildinObjects]
ON [dbo].[InboxDocuments]
    ([BuildingObj]);
GO

-- Creating foreign key on [BuildingObj] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_BuildinObjects]
    FOREIGN KEY ([BuildingObj])
    REFERENCES [dbo].[BuildinObjects]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_BuildinObjects'
CREATE INDEX [IX_FK_OutboxDocument_BuildinObjects]
ON [dbo].[OutboxDocuments]
    ([BuildingObj]);
GO

-- Creating foreign key on [Contractor] in table 'ContractorEmployes'
ALTER TABLE [dbo].[ContractorEmployes]
ADD CONSTRAINT [FK_ContractorEmployes_Contractors]
    FOREIGN KEY ([Contractor])
    REFERENCES [dbo].[Contractors]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_ContractorEmployes_Contractors'
CREATE INDEX [IX_FK_ContractorEmployes_Contractors]
ON [dbo].[ContractorEmployes]
    ([Contractor]);
GO

-- Creating foreign key on [Sender] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [FK_InboxDocuments_ContractorEmployes]
    FOREIGN KEY ([Sender])
    REFERENCES [dbo].[ContractorEmployes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_InboxDocuments_ContractorEmployes'
CREATE INDEX [IX_FK_InboxDocuments_ContractorEmployes]
ON [dbo].[InboxDocuments]
    ([Sender]);
GO

-- Creating foreign key on [RecieverEmploye] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_ContractorEmployes]
    FOREIGN KEY ([RecieverEmploye])
    REFERENCES [dbo].[ContractorEmployes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_ContractorEmployes'
CREATE INDEX [IX_FK_OutboxDocument_ContractorEmployes]
ON [dbo].[OutboxDocuments]
    ([RecieverEmploye]);
GO

-- Creating foreign key on [RecieverOrg] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_Contractors]
    FOREIGN KEY ([RecieverOrg])
    REFERENCES [dbo].[Contractors]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_Contractors'
CREATE INDEX [IX_FK_OutboxDocument_Contractors]
ON [dbo].[OutboxDocuments]
    ([RecieverOrg]);
GO

-- Creating foreign key on [DocState] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_DocStates]
    FOREIGN KEY ([DocState])
    REFERENCES [dbo].[DocStates]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_DocStates'
CREATE INDEX [IX_FK_OutboxDocument_DocStates]
ON [dbo].[OutboxDocuments]
    ([DocState]);
GO

-- Creating foreign key on [Reciever] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [FK_InboxDocuments_Employes]
    FOREIGN KEY ([Reciever])
    REFERENCES [dbo].[Employes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_InboxDocuments_Employes'
CREATE INDEX [IX_FK_InboxDocuments_Employes]
ON [dbo].[InboxDocuments]
    ([Reciever]);
GO

-- Creating foreign key on [WhoMade] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_EmployesWhoMade]
    FOREIGN KEY ([WhoMade])
    REFERENCES [dbo].[Employes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_EmployesWhoMade'
CREATE INDEX [IX_FK_OutboxDocument_EmployesWhoMade]
ON [dbo].[OutboxDocuments]
    ([WhoMade]);
GO

-- Creating foreign key on [WhoSign] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_EmployesWhoSign]
    FOREIGN KEY ([WhoSign])
    REFERENCES [dbo].[Employes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_EmployesWhoSign'
CREATE INDEX [IX_FK_OutboxDocument_EmployesWhoSign]
ON [dbo].[OutboxDocuments]
    ([WhoSign]);
GO

-- Creating foreign key on [Employe] in table 'Users'
ALTER TABLE [dbo].[Users]
ADD CONSTRAINT [FK_Users_Employes]
    FOREIGN KEY ([Employe])
    REFERENCES [dbo].[Employes]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_Users_Employes'
CREATE INDEX [IX_FK_Users_Employes]
ON [dbo].[Users]
    ([Employe]);
GO

-- Creating foreign key on [Files] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [FK_InboxDocuments_FileMetaData]
    FOREIGN KEY ([Files])
    REFERENCES [dbo].[FileMetaDatas]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_InboxDocuments_FileMetaData'
CREATE INDEX [IX_FK_InboxDocuments_FileMetaData]
ON [dbo].[InboxDocuments]
    ([Files]);
GO

-- Creating foreign key on [Files] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_FileMetaData]
    FOREIGN KEY ([Files])
    REFERENCES [dbo].[FileMetaDatas]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_FileMetaData'
CREATE INDEX [IX_FK_OutboxDocument_FileMetaData]
ON [dbo].[OutboxDocuments]
    ([Files]);
GO

-- Creating foreign key on [ResponseOn] in table 'InboxDocuments'
ALTER TABLE [dbo].[InboxDocuments]
ADD CONSTRAINT [FK_InboxDocs_OutboxDocument]
    FOREIGN KEY ([ResponseOn])
    REFERENCES [dbo].[OutboxDocuments]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_InboxDocs_OutboxDocument'
CREATE INDEX [IX_FK_InboxDocs_OutboxDocument]
ON [dbo].[InboxDocuments]
    ([ResponseOn]);
GO

-- Creating foreign key on [ResponseOn] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_InboxDocs]
    FOREIGN KEY ([ResponseOn])
    REFERENCES [dbo].[InboxDocuments]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_InboxDocs'
CREATE INDEX [IX_FK_OutboxDocument_InboxDocs]
ON [dbo].[OutboxDocuments]
    ([ResponseOn]);
GO

-- Creating foreign key on [TypeOfOutboxDoc] in table 'OutboxDocuments'
ALTER TABLE [dbo].[OutboxDocuments]
ADD CONSTRAINT [FK_OutboxDocument_TypesOfOutboxDocs]
    FOREIGN KEY ([TypeOfOutboxDoc])
    REFERENCES [dbo].[TypesOfOutboxDocs]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_OutboxDocument_TypesOfOutboxDocs'
CREATE INDEX [IX_FK_OutboxDocument_TypesOfOutboxDocs]
ON [dbo].[OutboxDocuments]
    ([TypeOfOutboxDoc]);
GO

-- Creating foreign key on [ParentDepartmentId] in table 'CompanyStructures'
ALTER TABLE [dbo].[CompanyStructures]
ADD CONSTRAINT [FK_CompanyStructure_CompanyStructure]
    FOREIGN KEY ([ParentDepartmentId])
    REFERENCES [dbo].[CompanyStructures]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_CompanyStructure_CompanyStructure'
CREATE INDEX [IX_FK_CompanyStructure_CompanyStructure]
ON [dbo].[CompanyStructures]
    ([ParentDepartmentId]);
GO

-- Creating foreign key on [DepartmentId] in table 'Employes'
ALTER TABLE [dbo].[Employes]
ADD CONSTRAINT [FK_Employes_CompanyStructure]
    FOREIGN KEY ([DepartmentId])
    REFERENCES [dbo].[CompanyStructures]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_Employes_CompanyStructure'
CREATE INDEX [IX_FK_Employes_CompanyStructure]
ON [dbo].[Employes]
    ([DepartmentId]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------