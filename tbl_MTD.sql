create table tbl_MTD
	(
		mtd_PaymentDatePosted datetime,
		mtd_SaleItemDatePosted datetime,
		mtd_SaleItemReference varchar(100),
		mtd_AttorneyDocket varchar(100),
		mtd_Status varchar(100),
		mtd_TransactionID varchar(100),
		mtd_SaleID varchar(100),
		mtd_Feecode varchar(100),
		mtd_FeeCodeDescription varchar(max),
		mtd_ItemPrice varchar(100),
		mtd_Quantity varchar(100),
		mtd_ItemTotal varchar(100),
		mtd_CustomerName varchar(100),
		mtd_FileName varchar(100),
		mtd_Importdate datetime,
		mtd_BatchID int
		)

create procedure sp_ImportMTD
		@mtd_PaymentDatePosted varchar(100),
		@mtd_SaleItemDatePosted varchar(100),
		@mtd_SaleItemReference varchar(100),
		@mtd_AttorneyDocket varchar(100),
		@mtd_Status varchar(100),
		@mtd_TransactionID varchar(100),
		@mtd_SaleID varchar(100),
		@mtd_Feecode varchar(100),
		@mtd_FeeCodeDescription varchar(max),
		@mtd_ItemPrice varchar(100),
		@mtd_Quantity varchar(100),
		@mtd_ItemTotal varchar(100),
		@mtd_CustomerName varchar(100),
		@mtd_FileName varchar(100)

AS
Begin
	declare @mts_ImportDate datetime
	set @mts_ImportDate = getdate()

	INSERT INTO tbl_MTD
			(

				mtd_PaymentDatePosted,
				mtd_SaleItemDatePosted,
				mtd_SaleItemReference,
				mtd_AttorneyDocket,
				mtd_Status,
				mtd_TransactionID,
				mtd_SaleID,
				mtd_Feecode,
				mtd_FeeCodeDescription,
				mtd_ItemPrice,
				mtd_Quantity,
				mtd_ItemTotal,
				mtd_CustomerName,
				mtd_FileName,
				mtd_Importdate,
				mtd_BatchID
			)
		VALUES
			(
				cast(@mtd_PaymentDatePosted as datetime),
				cast(@mtd_SaleItemDatePosted as datetime),
				RTRIM(LTRIM(@mtd_SaleItemReference)),
				RTRIM(LTRIM(@mtd_AttorneyDocket)),
				RTRIM(LTRIM(@mtd_Status)),
				RTRIM(LTRIM(@mtd_TransactionID)),
				RTRIM(LTRIM(@mtd_SaleID)),
				RTRIM(LTRIM(@mtd_Feecode)),
				RTRIM(LTRIM(@mtd_FeeCodeDescription)),
				RTRIM(LTRIM(@mtd_ItemPrice)),
				RTRIM(LTRIM(@mtd_Quantity)),
				RTRIM(LTRIM(@mtd_ItemTotal)),
				RTRIM(LTRIM(@mtd_CustomerName)),
				RTRIM(LTRIM(@mtd_FileName)),
				RTRIM(LTRIM(@mtd_Importdate)),
				NULL
			)
End

Select COUNT(*) As RecCount
From tbl_MTS