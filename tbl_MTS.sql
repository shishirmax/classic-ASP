create table tbl_MTS
	(
		mts_CsvRefNumber int identity(1,1),
		mts_DatePosted datetime,
		mts_Transactionreference varchar(100),
		mts_AttorneyDocket varchar(100),
		mts_Status varchar(100),
		mts_TransactionID varchar(100),
		mts_Type varchar(100),
		mts_TotalPayment_Refund varchar(100),
		mts_CustomerName varchar(100),
		mts_FileName varchar(100),
		mts_ImportDate datetime,
		mts_BatchID int
	)


create procedure sp_ImportMTS
	@mts_DatePosted varchar(100),
	@mts_Transactionreference varchar(100),
	@mts_AttorneyDocket varchar(100),
	@mts_Status varchar(100),
	@mts_TransactionID varchar(100),
	@mts_Type varchar(100),
	@mts_TotalPayment_Refund varchar(100),
	@mts_CustomerName varchar(100),
	@mts_FileName varchar(100)

AS
Begin
	declare @mts_ImportDate datetime
	set @mts_ImportDate = getdate()

		INSERT INTO tbl_MTS
				(
					mts_DatePosted,
					mts_Transactionreference,
					mts_AttorneyDocket,
					mts_Status,
					mts_TransactionID,
					mts_Type,
					mts_TotalPayment_Refund,
					mts_CustomerName,
					mts_FileName,
					mts_ImportDate,
					mts_BatchID
				)
		VALUES
				(
					CAST(@mts_DatePosted as DATETIME),
					RTRIM(LTRIM(@mts_Transactionreference)),
					RTRIM(LTRIM(@mts_AttorneyDocket)),
					RTRIM(LTRIM(@mts_Status)),
					RTRIM(LTRIM(@mts_TransactionID)),
					RTRIM(LTRIM(@mts_Type)),
					RTRIM(LTRIM(@mts_TotalPayment_Refund)),
					RTRIM(LTRIM(@mts_CustomerName)),
					RTRIM(LTRIM(@mts_FileName)),
					RTRIM(LTRIM(@mts_ImportDate)),
					NULL
				)
End

Select COUNT(*) As RecCount
From tbl_MTS