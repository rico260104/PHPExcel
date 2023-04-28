<?php
defined('BASEPATH') or exit('No direct script access allowed');

class C_cetak_faktur extends CI_Controller
{
    public function __construct()
    {
        parent::__construct();
        $this->ci = &get_instance();
        $this->load->library('session');
        $this->load->library(array('PHPExcel_/PHPExcel', 'PHPExcel_/PHPExcel/IOFactory'));
        $this->load->helper('url');
        $this->load->helper('form');
        $this->load->model('m_customer');
        $this->load->model('m_api');
        $this->load->model('m_api_mo');
        $this->load->model('m_master');
        $this->load->model('m_faktur');


        ini_set('max_execution_time', 0);
        ini_set('memory_limit', '2048M');
    }
    public function index()
    {
        //membuat objek PHPExcel
        $objPHPExcel = new PHPExcel();

        //Start adding next sheets
        $namesheet = array('List Data Pajak');
        //jumlah header
        $jumlah = array(21);

        $header = array(
            array(
                'FK',
                'KD_JNS_TRANSAKSI',
                'FG_PENGGANTI',
                'NOMOR_FAKTUR',
                'MASA_PAJAK',
                'TAHUN_PAJAK',
                'TANGGAL_FAKTUR',
                'NPWP',
                'NAMA',
                'ALAMAT_LENGKAP',
                'JUMLAH_DPP',
                'JUMLAH_PPN',
                'JUMLAH_PPNBM',
                'ID_KETERANGAN_TAMBAHAN',
                'FG_UANG_MUKA',
                'UANG_MUKA_DPP',
                'UANG_MUKA_PPN',
                'UANG_MUKA_PPNBM',
                'REFERENSI',
                'KODE_DOKUMEN_PENDUKUNG',
                'Column1',
            )

        );
        $created_user = $this->session->userdata('user');
        // $area = $this->session->userdata('AREADISTRIBUTOR');
        //memisahkan multiple choice
        $tgl = explode("-", $this->input->post('tgl'));
        $tgl1 = str_replace("/", "-", $tgl[0]);
        $tgl2 = str_replace("/", "-", $tgl[1]);
        $datas = $this->m_faktur
            ->get_data_fk($tgl1, $tgl2, $created_user);
        $datas2 = $this->m_faktur
            ->get_data_of($tgl1, $tgl2, $created_user);
        $datas3 = $this->m_faktur
            ->get_data_lt($tgl1, $tgl2, $created_user);
        //========================== Tambah sheet dan beri header tabel ==========================
        $i = 0;
        while ($i < 1) {
            // Add new sheet
            $objWorkSheet = $objPHPExcel->createSheet($i); //Setting index when creating                
            $lastColumn = $objWorkSheet->getHighestColumn();

            for ($x = 0; $x < $jumlah[$i]; $x++) {
                $row = 1;
                $objWorkSheet->setCellValue($lastColumn . $row, $header[$i][$x]);
                $objWorkSheet
                    ->getColumnDimension($lastColumn)
                    ->setAutoSize(true);
                $objWorkSheet->getStyle($lastColumn . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                $objWorkSheet->getStyle($lastColumn . $row)->getFill()->getStartColor()->setARGB('70ad47');
                $lastColumn++;
            }


            $objWorkSheet->setTitle($namesheet[$i]);
            $i++;
        }


        //============================== Tulis data dari query ===================================
        $index = 0;
        $rowgabung = 2;
        while ($index < 1) {
            $worksheet = $objPHPExcel->setActiveSheetIndex($index);

            $trigger = 0;
            $row = 4;
            $rowlt = 2;
            $rowof = 3;
            if ($index == 0) {
                // header lt
                $worksheet->setCellValue('A' . $rowlt, 'LT');
                $worksheet->setCellValue('B' . $rowlt, 'NPWP');
                $worksheet->setCellValue('C' . $rowlt, 'NAMA');
                $worksheet->setCellValue('D' . $rowlt, 'JALAN');
                $worksheet->setCellValue('E' . $rowlt, 'BLOK');
                $worksheet->setCellValue('F' . $rowlt, 'NOMOR');
                $worksheet->setCellValue('G' . $rowlt, 'RT');
                $worksheet->setCellValue('H' . $rowlt, 'RW');
                $worksheet->setCellValue('I' . $rowlt, 'KECAMATAN');
                $worksheet->setCellValue('J' . $rowlt, 'KELURAHAN');
                $worksheet->setCellValue('K' . $rowlt, 'KABUPATEN');
                $worksheet->setCellValue('L' . $rowlt, 'PROVINSI');
                $worksheet->setCellValue('M' . $rowlt, 'KODE_POS');
                $worksheet->setCellValue('N' . $rowlt, 'NOMOR_TELEPON');
                // header of
                $worksheet->setCellValue('A' . $rowof, 'OF');
                $worksheet->setCellValue('B' . $rowof, 'KODE_OBJEK');
                $worksheet->setCellValue('C' . $rowof, 'NAMA');
                $worksheet->setCellValue('D' . $rowof, 'HARGA_SATUAN');
                $worksheet->setCellValue('E' . $rowof, 'JUMLAH_BARANG');
                $worksheet->setCellValue('F' . $rowof, 'HARGA_TOTAL');
                $worksheet->setCellValue('G' . $rowof, 'DISKON');
                $worksheet->setCellValue('H' . $rowof, 'DPP');
                $worksheet->setCellValue('I' . $rowof, 'PPN');
                $worksheet->setCellValue('J' . $rowof, 'TARIF_PPNBM');
                $worksheet->setCellValue('K' . $rowof, 'PPN');


                foreach ($datas as $data) {
                    //    untuk data fk
                    $kode = 'FK';
                    //Kolom a
                    $worksheet->setCellValue('A' . $row, $kode);
                    $worksheet->getStyle('A' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('A' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom b
                    $worksheet->setCellValueExplicit('B' . $row, '01', PHPExcel_Cell_DataType::TYPE_STRING);
                    $worksheet->getStyle('B' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('B' . $row)->getFill()->getStartColor()->setARGB('70ad47');

                    //Kolom C
                    $worksheet->setCellValue('C' . $row, '0');
                    $worksheet->getStyle('C' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('C' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom D
                    $worksheet->setCellValue('D' . $row, $data['no_invoice']);
                    $worksheet->getStyle('D' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('D' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom E
                    $worksheet->setCellValue('E' . $row, $data['masa_pajak']);
                    $worksheet->getStyle('E' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('E' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom F
                    $worksheet->setCellValue('F' . $row, $data['tahun_pajak']);
                    $worksheet->getStyle('F' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('F' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom G
                    $worksheet->setCellValue('G' . $row, $data['tanggal_invoice']);
                    $worksheet->getStyle('G' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('G' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom H
                    $worksheet->setCellValue('H' . $row, $data['npwp']);
                    $worksheet->getStyle('H' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('H' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom I

                    if ($data['npwp'] == '') {
                        $worksheet->setCellValue('I' . $row, $data['nik'] . '#NIK#NAMA#' . $data['nama_toko']);
                    } else {
                        $worksheet->setCellValue('I' . $row, $data['nama_toko']);
                    }

                    $worksheet->getStyle('I' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('I' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom J
                    $worksheet->setCellValue('J' . $row, $data['alamat']);
                    $worksheet->getStyle('J' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('J' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom K
                    $worksheet->setCellValue('K' . $row, $data['total_invoice']);
                    $worksheet->getStyle('K' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('K' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom L
                    $worksheet->setCellValue('L' . $row, $data['total_invoice'] * 11 / 100);
                    $worksheet->getStyle('L' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('L' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom M
                    $worksheet->setCellValue('M' . $row, '0');
                    $worksheet->getStyle('M' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('M' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom N
                    $worksheet->setCellValue('N' . $row, '-');
                    $worksheet->getStyle('N' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('N' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom O
                    $worksheet->setCellValue('O' . $row, '0');
                    $worksheet->getStyle('O' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('O' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom P
                    $worksheet->setCellValue('P' . $row, '0');
                    $worksheet->getStyle('P' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('P' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom Q
                    $worksheet->setCellValue('Q' . $row, '0');
                    $worksheet->getStyle('Q' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('Q' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom R
                    $worksheet->setCellValue('R' . $row, '0');
                    $worksheet->getStyle('R' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('R' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom S
                    $worksheet->setCellValue('S' . $row, 'Nomer #: ' . $data['no_invoice'] . 'NIK: ' . $data['nik']);
                    $worksheet->getStyle('S' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('S' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom T
                    $worksheet->setCellValue('T' . $row, '');
                    $worksheet->getStyle('T' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('T' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom U
                    $worksheet->setCellValue('U' . $row, '');
                    $worksheet->getStyle('U' . $row)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('U' . $row)->getFill()->getStartColor()->setARGB('70ad47');
                    //Kolom a

                    $row++;
                    // untuk data Lt       

                    foreach ($datas3 as $datalt) {
                        if ($datalt['npwp'] == $data['npwp']) {
                            if ($datalt['npwp'] != '') {
                                $kode = 'LT';
                                $worksheet->setCellValue('A' . $row, $kode);
                                $worksheet->setCellValue('B' . $row, $datalt['npwp']);
                                $worksheet->setCellValue('C' . $row, $datalt['nama_toko']);
                                $worksheet->setCellValue('D' . $row, $datalt['alamat']);
                                $worksheet->setCellValue('E' . $row, '-');
                                $worksheet->setCellValue('F' . $row, '-');
                                $worksheet->setCellValue('G' . $row, '-');
                                $worksheet->setCellValue('H' . $row, '-');
                                $worksheet->setCellValue('I' . $row, '-');
                                $worksheet->setCellValue('J' . $row, '-');
                                $worksheet->setCellValue('K' . $row, '-');
                                $worksheet->setCellValue('L' . $row, '-');
                                $worksheet->setCellValue('M' . $row, '-');
                                $worksheet->setCellValue('N' . $row, $datalt['no_hp']);
                                $trigger++;
                                $row++;
                            }
                        }
                    }



                    // untuk data OF
                    foreach ($datas2 as $dataof) {
                        if ($dataof['no_invoice'] == $data['no_invoice']) {
                            $kode = 'OF';
                            $worksheet->setCellValue('A' . $row, $kode);
                            $worksheet->setCellValue('B' . $row, $dataof['kode_barang']);
                            $worksheet->setCellValue('C' . $row, $dataof['nama']);
                            $worksheet->setCellValue('D' . $row, $dataof['harga']);
                            $worksheet->setCellValue('E' . $row, $dataof['jumlah']);
                            $worksheet->setCellValue('F' . $row, $dataof['harga'] * $dataof['jumlah']);
                            $worksheet->setCellValue('G' . $row, '0');
                            $worksheet->setCellValue('H' . $row, $dataof['harga'] * $dataof['jumlah']);
                            $worksheet->setCellValue('I' . $row, $dataof['harga'] * $dataof['jumlah'] * 11 / 100);
                            $worksheet->setCellValue('J' . $row, '0');
                            $worksheet->setCellValue('K' . $row, '0');


                            $row++;
                        }
                    }
                }
            }

            $index++;
        }

        //mulai menyimpan excel format xlsx, kalau ingin xls ganti Excel2007 menjadi Excel5          
        //$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter = IOFactory::createWriter($objPHPExcel, 'Excel2007');
        ob_end_clean();
        //sesuaikan headernya 
        header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
        header("Cache-Control: no-store, no-cache, must-revalidate");
        header("Cache-Control: post-check=0, pre-check=0", false);
        header("Pragma: no-cache");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        //ubah nama file saat diunduh
        header('Content-Disposition: attachment;filename="EKS FK.xlsx"');
        //unduh file
        $objWriter->save("php://output");
    }
}
