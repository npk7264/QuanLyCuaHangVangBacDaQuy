CREATE DATABASE VangBacDaQuy

USE VangBacDaQuy

set dateformat dmy

CREATE TABLE KHACHHANG
(
	MaKH varchar(5) not null,
	TenKH nvarchar(40), 
	SDT varchar(11),
	DiaChi nvarchar(50)
	constraint PK_KHACHHANG primary key(MaKH)
)

CREATE TABLE NHANVIEN
(
	MaNV varchar(5) not null,
	TenNV nvarchar(40),
	GioiTinh nvarchar(5)
	constraint PK_NHANVIEN primary key(MaNV)
)

CREATE TABLE PHIEUBANHANG
(
	SoPhieu varchar(30) not null,
	NgayLap smalldatetime,
	MaKH varchar(5) not null,
	MaNV varchar(5) not null,
	TongTien float
	constraint PK_PHIEUBANHANG primary key(SoPhieu)
)

CREATE TABLE CTPBH
(
	SoPhieu varchar(30) not null,
	MaSP varchar(5) not null,
	SoLuong int,
	DonGiaBan float,
	GiamGia float,
	ThanhTien float
	constraint PK_CTPBH primary key(SoPhieu,MaSP)
)

CREATE TABLE SANPHAM
(
	MaSP varchar(5) not null,
	TenSP nvarchar(40),
	MaLoaiSP varchar(5) not null,
	SoLuong int,
	DonGiaMua float,
	DonGiaBan float
	constraint PK_SANPHAM primary key(MaSP)
)

CREATE TABLE LOAISANPHAM
(
	MaLoaiSP varchar(5) not null,
	TenLoaiSP nvarchar(30),
	MaDVT varchar(5) not null,
	PhanTramLoiNhuan float
	constraint PK_LOAISANPHAM primary key(MaLoaiSP)
)

CREATE TABLE DONVITINH
(
	MaDVT varchar(5) not null,
	TenDVT nvarchar(20),
    constraint PK_DONVITINH primary key(MaDVT)
)

CREATE TABLE NHACUNGCAP
(
	MaNCC varchar(5) not null,
	TenNCC nvarchar(30),
	DiaChi nvarchar(50),
	SDT varchar(11),
	constraint PK_NHACUNGCAP primary key(MaNCC)
)

CREATE TABLE PHIEUMUAHANG
(
	SoPhieu varchar(30) not null,
	NgayLap smalldatetime,
	MaNCC varchar(5) not null,
	MaNV varchar(5) not null,
	TongTien float
	constraint PK_PHIEUMUAHANG primary key(SoPhieu)
)

CREATE TABLE CTPMH
(
	SoPhieu varchar(30) not null,
	MaSP varchar(5) not null,
	SoLuong int,
	DonGiaMua float,
	ThanhTien float
	constraint PK_CTPMH primary key(SoPhieu, MaSP)
)

CREATE TABLE DICHVU
(
	MaDV varchar(5) not null,
	LoaiDV nvarchar(20) not null,
	DonGiaDV float,
	constraint PK_DICHVU primary key(MaDV)
)

CREATE TABLE PHIEUDICHVU
(
	SoPhieu varchar(30) not null,
	NgayLap smalldatetime,
	MaKH varchar(5) not null,
	MaNV varchar(5) not null,
	TongTien float,
	TongTienTraTruoc float,
	TongTienConLai float,
	TinhTrang nvarchar(20),
	constraint PK_PHIEUDICHVU primary key(SoPhieu)
)


CREATE TABLE CTPDV 
(
	SoPhieu varchar(30) not null,
	MaDV varchar(5) not null,
	ChiPhiRieng float,
	DonGiaDuocTinh float,
	SoLuong int,
	ThanhTien float,
	TraTruoc float,
	ConLai float,
	NgayGiao smalldatetime,
	TinhTrang nvarchar(10)
	constraint PK_CTPDV primary key(SoPhieu,MaDV)
)

CREATE TABLE CTBCTK
(
	Thang int not null,
	Nam int not null,
	MaSP varchar(5) not null,
	TonDau int,
	SoLuongMua int,
	SoLuongBan int,
	TonCuoi int,
	constraint PK_CTBCTK primary key(Thang, Nam, MaSP)
)

CREATE TABLE THAMSO
(
	TenThamSo varchar(20) not null,
	GiaTri int,
	constraint PK_THAMSO primary key(TenThamSo)
)


alter table PHIEUBANHANG add constraint FK_PHIEUBANHANG_KHACHHANG
foreign key (MaKH) references KHACHHANG(MaKH)
alter table PHIEUBANHANG add constraint FK_PHIEUBANHANG_NHANVIEN
foreign key (MaNV) references NHANVIEN(MaNV)

alter table CTPBH add constraint FK_CTPBH_SANPHAM 
foreign key(MaSP) references SANPHAM(MaSP)
alter table CTPBH add constraint FK_CTPBH_PHIEUBANHANG
foreign key(SoPhieu) references PHIEUBANHANG(SoPhieu)

alter table SANPHAM add constraint FK_SANPHAM_LOAISANPHAM
foreign key (MaLoaiSP) references LOAISANPHAM(MaLoaiSP)

alter table LOAISANPHAM add constraint FK_LOAISANPHAM_DONVITINH
foreign key (MaDVT) references DONVITINH(MaDVT)

alter table PHIEUMUAHANG add constraint FK_PHIEUMUAHANG_NHACUNGCAP
foreign key (MaNCC) references NHACUNGCAP(MaNCC)
alter table PHIEUMUAHANG add constraint FK_PHIEUMUAHANG_NHANVIEN
foreign key (MaNV) references NHANVIEN(MaNV)

alter table CTPMH add constraint FK_CTPMH_SANPHAM
foreign key (MaSP) references SANPHAM(MaSP)
alter table CTPMH add constraint FK_CTPMH_PHIEUMUAHANG
foreign key (SoPhieu) references PHIEUMUAHANG(SoPhieu)

alter table PHIEUDICHVU add constraint FK_PHIEUDICHVU_KHACHHANG
foreign key (MaKH) references KHACHHANG(MaKH)
alter table PHIEUDICHVU add constraint FK_PHIEUDICHVU_NHANVIEN
foreign key (MaNV) references NHANVIEN(MaNV)

alter table CTPDV add constraint FK_CTPDV_DICHVU
foreign key (MaDV) references DICHVU(MaDV)
alter table CTPDV add constraint FK_CTPDV_PHIEUDICHVU
foreign key (SoPhieu) references PHIEUDICHVU(SoPhieu)

alter table CTBCTK add constraint FK_CTBCTK_SANPHAM
foreign key (MaSP) references SANPHAM(MaSP)

INSERT INTO DONVITINH VALUES (N'DVT01', N'cara')
INSERT INTO DONVITINH VALUES (N'DVT02', N'gam')

INSERT INTO LOAISANPHAM VALUES (N'LSP01', N'Vàng 24', N'DVT01', 15)
INSERT INTO LOAISANPHAM VALUES (N'LSP02', N'Vàng 18', N'DVT01', 10)
INSERT INTO LOAISANPHAM VALUES (N'LSP03', N'Kim cương', N'DVT01', 20)

INSERT INTO SANPHAM VALUES (N'SP001', N'Dây chuyền vàng', N'LSP01', 0, 0, 0)
INSERT INTO SANPHAM VALUES (N'SP002', N'Dây chuyền vàng', N'LSP02', 0, 0, 0)

INSERT INTO DICHVU VALUES (N'DV001', N'Đánh bóng trang sức', 500000)
INSERT INTO DICHVU VALUES (N'DV002', N'Mạ vàng', 700000)

INSERT INTO KHACHHANG VALUES (N'KH001', N'Bùi Duy Anh Đức', N'01010101010', N'Quảng Trị')
INSERT INTO KHACHHANG VALUES (N'KH002', N'Nguyễn Phúc Khang', N'02020202020', N'An Giang')
INSERT INTO KHACHHANG VALUES (N'KH003', N'Võ Trung Kiên', N'03030303030', N'TPHCM')
INSERT INTO KHACHHANG VALUES (N'KH004', N'Hà Văn Linh', N'04040404040', N'Đak Lak')

INSERT INTO NHANVIEN VALUES (N'NV001', N'Hà Văn Linh', N'Nam')
INSERT INTO NHANVIEN VALUES (N'NV002', N'Võ Trung Kiên', N'Nam')
INSERT INTO NHANVIEN VALUES (N'NV003', N'Nguyễn Phúc Khang', N'Nam')
INSERT INTO NHANVIEN VALUES (N'NV004', N'Bùi Duy Anh Đức', N'Nam')

INSERT INTO NHACUNGCAP VALUES (N'NCC01', N'Công ty vàng Kim Hương', N'An Giang', N'01234567899')
INSERT INTO NHACUNGCAP VALUES (N'NCC02', N'Đại lý trang sức ', N'TPHCM', N'06574321456')





