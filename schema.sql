-- ====================================================================
-- PROJECT: Centralized Attendance Automation System
-- PURPOSE: Database schema for managing branch logins and daily attendance
-- ====================================================================

-- Table 1: Branchlogin
-- Stores secure credentials (UserID & Password) and respective branch names 
-- for each school branch to ensure only authorized staff can log in.
Create Table Branchlogin(
    SNO int identity(1,1) primary key,
    UserID varchar(30) not null,
    Password varchar(30) not null,
    Branch varchar(30) not null
);

-- Table 2: Attendance
-- Main transactional table storing daily class-wise standard and present 
-- student counts submitted by each branch. 
-- Features a UNIQUE constraint (`BranchName, EntryDate`) to prevent 
-- duplicate submissions by the same branch on the same day.
Create Table Attendance(
	ID int identity(1,1) primary key
	,EntryDate Date default cast(Getdate() as date)
	,BranchName varchar(50) not null
	,Class1_Standard int not null
	,Class1_Present int not null
	,Class2_Standard int not null
	,Class2_Present int not null
	,Class3_Standard int not null
	,Class3_Present int not null
	,Class4_Standard int not null
	,Class4_Present int not null
	,Class5_Standard int not null
	,Class5_Present int not null
	,Class6_Standard int not null
	,Class6_Present int not null
	,Class7_Standard int not null
	,Class7_Present int not null
	,Class8_Standard int not null
	,Class8_Present int not null
	,Class9_Standard int not null
	,Class9_Present int not null
	,Class10_Standard int not null
	,Class10_Present int not null
	,Constraint UN_Branch_Date UNIQUE(BranchName,EntryDate)
);
