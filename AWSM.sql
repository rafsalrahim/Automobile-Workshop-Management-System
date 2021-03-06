USE [AWSM]
GO
/****** Object:  Table [dbo].[tbl_wstatus]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_wstatus](
	[ws] [int] IDENTITY(1,1) NOT NULL,
	[u_id] [int] NULL,
	[name] [varchar](50) NULL,
	[vh_id] [int] NULL,
	[f_date] [date] NULL,
	[status] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_wheelrate]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_wheelrate](
	[tr_id] [int] IDENTITY(1,1) NOT NULL,
	[service] [varchar](50) NULL,
	[rate] [bigint] NULL,
	[vh_company] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_wheelbill]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_wheelbill](
	[wh_id] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NOT NULL,
	[amt] [bigint] NULL,
	[model] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_wheel]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_wheel](
	[t_id] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[vh_compnam] [varchar](50) NULL,
	[vh_model] [varchar](50) NULL,
	[t_type] [varchar](50) NULL,
	[f_date] [datetime] NULL,
	[t_date] [datetime] NULL,
	[status] [varchar](50) NULL,
	[regi_no] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_waterrate]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_waterrate](
	[wr_id] [int] IDENTITY(1,1) NOT NULL,
	[catagory] [varchar](50) NULL,
	[rate] [bigint] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_watere]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_watere](
	[e] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[fulbdy] [varchar](50) NULL,
	[chase] [varchar](50) NULL,
	[Others] [varchar](50) NULL,
	[tot_amt] [int] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_water]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_water](
	[w_id] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[vh_compnam] [varchar](50) NULL,
	[vh_model] [varchar](50) NULL,
	[fulbdy] [varchar](max) NULL,
	[chase] [varchar](max) NULL,
	[Others] [varchar](max) NULL,
	[f_date] [datetime] NULL,
	[t_date] [datetime] NULL,
	[status] [varchar](50) NULL,
	[catagory] [varchar](50) NULL,
	[regi_no] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_vehicleregistration]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_vehicleregistration](
	[vh_id] [int] IDENTITY(1,1) NOT NULL,
	[name] [varchar](50) NULL,
	[address] [varchar](max) NULL,
	[mob_no] [bigint] NULL,
	[email] [varchar](max) NULL,
	[vh_compnam] [varchar](max) NULL,
	[vh_model] [varchar](max) NULL,
	[regi_no] [varchar](max) NULL,
	[chai_no] [varchar](max) NULL,
	[date] [date] NULL,
	[status] [varchar](max) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_user]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_user](
	[u_id] [int] IDENTITY(1,1) NOT NULL,
	[u_name] [varchar](50) NULL,
	[u_add] [varchar](50) NULL,
	[u_cont] [bigint] NULL,
	[usernam] [varchar](50) NULL,
	[password] [varchar](50) NULL,
	[usertype] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_spareparts]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_spareparts](
	[s_id] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[vh_compnam] [varchar](50) NULL,
	[vh_model] [varchar](50) NULL,
	[partnam] [varchar](50) NULL,
	[discription] [varchar](max) NULL,
	[status] [varchar](10) NULL,
	[price] [bigint] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_reprate]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_reprate](
	[r_id] [int] IDENTITY(1,1) NOT NULL,
	[type_ser] [varchar](50) NULL,
	[rate] [bigint] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_repbill]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_repbill](
	[w] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[engine] [varchar](50) NULL,
	[brake] [varchar](50) NULL,
	[oil] [varchar](50) NULL,
	[others] [varchar](50) NULL,
	[tot_amt] [int] NULL,
	[bas_amt] [int] NULL,
	[spr_amt] [int] NULL,
	[spare] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_repare]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_repare](
	[rid] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[engine] [varchar](30) NULL,
	[brake] [varchar](30) NULL,
	[oil] [varchar](30) NULL,
	[others] [varchar](30) NULL,
	[discription] [varchar](max) NULL,
	[frm_date] [datetime] NULL,
	[to_date] [datetime] NULL,
	[status] [varchar](20) NULL,
	[vh_model] [varchar](50) NULL,
	[vh_company] [varchar](50) NULL,
	[regi_no] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_login]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_login](
	[login_id] [int] IDENTITY(1,1) NOT NULL,
	[username] [varchar](20) NULL,
	[password] [varchar](30) NULL,
	[usr_type] [varchar](20) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_bodywork]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_bodywork](
	[vh_id] [int] NOT NULL,
	[cust_id2] [int] IDENTITY(1,1) NOT NULL,
	[paint] [varchar](50) NULL,
	[bdywork] [varchar](50) NULL,
	[glass] [varchar](max) NULL,
	[others] [varchar](30) NULL,
	[description] [varchar](50) NULL,
	[fdate] [datetime] NULL,
	[tdate] [datetime] NULL,
	[status] [varchar](20) NULL,
	[vh_compnam] [varchar](50) NULL,
	[vh_model] [varchar](50) NULL,
	[regi_no] [varchar](50) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_bdwrkrate]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_bdwrkrate](
	[b_id] [int] IDENTITY(1,1) NOT NULL,
	[type] [varchar](max) NULL,
	[rate] [bigint] NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tbl_bdbill]    Script Date: 01/15/2016 19:04:59 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_bdbill](
	[q] [int] IDENTITY(1,1) NOT NULL,
	[vh_id] [int] NULL,
	[bdywork] [varchar](max) NULL,
	[paint] [varchar](max) NULL,
	[glass] [varchar](max) NULL,
	[tot_amt] [int] NULL,
	[spare] [varchar](max) NULL,
	[bas_amt] [int] NULL,
	[spr_amt] [int] NULL,
	[others] [varchar](max) NULL
) ON [PRIMARY]
GO
SET ANSI_PADDING OFF
GO
