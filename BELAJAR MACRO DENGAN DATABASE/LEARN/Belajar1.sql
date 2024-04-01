CREATE TABLE tb_penumpang(
	Nama_Depan VARCHAR(255) NOT NULL,
	Nama_Belakang VARCHAR(255) NOT NULL,
	Stasiun_Keberangkatan VARCHAR(255) NOT NULL,
	Kedatangan VARCHAR(255) NOT NULL,
	Boarding VARCHAR(255) NOT NULL,
	Sampai VARCHAR(255) NOT NULL
);

DROP TABLE db_dicoding.`tb_penumpang`

DESCRIBE db_dicoding.`tb_penumpang`

ALTER TABLE db_dicoding.`tb_penumpang`
MODIFY Boarding VARCHAR(255) NOT NULL, 
MODIFY Sampai VARCHAR(255) NOT NULL;

SELECT * FROM `tb_penumpang`

/*Add Values*/
INSERT INTO `tb_penumpang` VALUES
('Defanty','Veninda','Cimahi','Jakarta','18.30','21.20'),
('Galuh','Suparman','Cimahi','Jakarta','18.30','21.20'),
('Giantinisa','Salma','Bandung','Jakarta','18.30','21.20'),
('Hanifa','Supartiwi','Bekasi','Jakarta','18.30','21.20'),
('Maria ','Rizma','Bandung','Jakarta','18.30','21.20'),
('Sri','Ayu','Bandung','Jakarta','18.30','21.20'),
('Yunita','Priatna','Bandung','Jakarta','18.30','21.20')
;
`tb_penumpang`
/*Uniq Values*/
SELECT DISTINCT
    Nama_Depan,
    Nama_Belakang,
    Stasiun_Keberangkatan,
    Kedatangan,
    Boarding,
    Sampai
FROM
    tb_penumpang;

SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'db_dicoding.tb_penumpang'