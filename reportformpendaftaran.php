<?php
//deklarasi variabel konfigurasi database
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "latihan2";

//deklarasi varibel untuk koneksi database
$koneksi = mysqli_connect($servername, $username, $password, $dbname);

//cek Koneksi
if (!$koneksi) {
  die("Connection failed: " . mysqli_connect_error());
}

//memanggil library
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//menuliskan nama kolom pada excel
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Jenis Pendaftaran');
$sheet->setCellValue('B1', 'Tanggal Masuk');
$sheet->setCellValue('C1', 'NIS');
$sheet->setCellValue('D1', 'Nomor Peserta Ujian');
$sheet->setCellValue('E1', 'Pernah Paud?');
$sheet->setCellValue('F1', 'Pernah TK?');
$sheet->setCellValue('G1', 'No SKHUN Sebelumnya');
$sheet->setCellValue('H1', 'No Ijazah Sebelumnya');
$sheet->setCellValue('I1', 'Hobi');
$sheet->setCellValue('J1', 'Cita-Cita');
$sheet->setCellValue('K1', 'Nama Lengkap');
$sheet->setCellValue('L1', 'Jenis Kelamin');
$sheet->setCellValue('M1', 'No NISN');
$sheet->setCellValue('N1', 'No NIK');
$sheet->setCellValue('O1', 'Tempat Lahir');
$sheet->setCellValue('P1', 'Tanggal Lahir');
$sheet->setCellValue('Q1', 'Agama');
$sheet->setCellValue('R1', 'Berkebutuhan Khusus');
$sheet->setCellValue('S1', 'Alamat');
$sheet->setCellValue('T1', 'RT');
$sheet->setCellValue('U1', 'RW');
$sheet->setCellValue('V1', 'Nama Dusun');
$sheet->setCellValue('W1', 'Nama Kelurahan/Desa');
$sheet->setCellValue('X1', 'Nama Kecamatan');
$sheet->setCellValue('Y1', 'Kode Pos');
$sheet->setCellValue('Z1', 'Tempat Tinggal');
$sheet->setCellValue('AA1', 'Moda Transportasi');
$sheet->setCellValue('AB1', 'No HP');
$sheet->setCellValue('AC1', 'No Telp');
$sheet->setCellValue('AD1', 'Email Pribadi');
$sheet->setCellValue('AE1', 'Penerima KPS/PKH/KIP');
$sheet->setCellValue('AF1', 'No KPS/PKH/KIP');
$sheet->setCellValue('AG1', 'Kewarganegaraan');

//mengambil data pada database dan menuliskan pada excel
$query = mysqli_query($koneksi,"select * from pendaftaran");
$i = 2;
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $row['jenis_pendaftaran']);
	$sheet->setCellValue('B'.$i, $row['tanggal_masuk']);
	$sheet->setCellValue('C'.$i, $row['nis']);
	$sheet->setCellValue('D'.$i, $row['nomor_peserta']);
	$sheet->setCellValue('E'.$i, $row['paud']);
	$sheet->setCellValue('F'.$i, $row['tk']);
	$sheet->setCellValue('G'.$i, $row['no_skhun']);
	$sheet->setCellValue('H'.$i, $row['no_ijazah']);
	$sheet->setCellValue('I'.$i, $row['hobi']);
	$sheet->setCellValue('J'.$i, $row['cita_cita']);
	$sheet->setCellValue('K'.$i, $row['jenis_kelamin']);
	$sheet->setCellValue('L'.$i, $row['nama']);
	$sheet->setCellValue('M'.$i, $row['nisn']);
	$sheet->setCellValue('N'.$i, $row['nik']);
	$sheet->setCellValue('O'.$i, $row['tempat_lahir']);
	$sheet->setCellValue('P'.$i, $row['tanggal_lahir']);
	$sheet->setCellValue('Q'.$i, $row['agama']);
	$sheet->setCellValue('R'.$i, $row['berkebutuhan_khusus']);
	$sheet->setCellValue('S'.$i, $row['alamat']);
	$sheet->setCellValue('T'.$i, $row['rt']);
	$sheet->setCellValue('U'.$i, $row['rw']);
	$sheet->setCellValue('V'.$i, $row['dusun']);
	$sheet->setCellValue('W'.$i, $row['kelurahan']);
	$sheet->setCellValue('X'.$i, $row['kecamatan']);
	$sheet->setCellValue('Y'.$i, $row['kode_pos']);
	$sheet->setCellValue('Z'.$i, $row['tempat_tinggal']);
	$sheet->setCellValue('AA'.$i, $row['transportasi']);
	$sheet->setCellValue('AB'.$i, $row['no_hp']);
	$sheet->setCellValue('AC'.$i, $row['no_telp']);
	$sheet->setCellValue('AD'.$i, $row['email']);
	$sheet->setCellValue('AE'.$i, $row['penerima_kps']);
	$sheet->setCellValue('AF'.$i, $row['no_kps']);
	$sheet->setCellValue('AG'.$i, $row['kewarganegaraan']);
	$i++;
}

//style
$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:Y'.$i)->applyFromArray($styleArray);

//memunculkan file excel
$writer = new Xlsx($spreadsheet);
$writer->save('Report Pendaftaran Siswa Baru.xlsx');
?>
