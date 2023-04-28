<?php
defined('BASEPATH') or exit('No direct script access allowed');

class C_rekap_kunjungan extends CI_Controller
{
    public function __construct()
    {
        parent::__construct();
        $this->ci = &get_instance();
        $this->load->library('session');
        $this->load->library(array('PHPExcel_/PHPExcel', 'PHPExcel_/PHPExcel/IOFactory'));
        $this->load->helper('url');
        $this->load->helper('form');
        $this->load->model('m_kunjungan');



        ini_set('max_execution_time', 0);
        ini_set('memory_limit', '2048M');
    }
    public function index()
    {
        //membuat objek PHPExcel
        $objPHPExcel = new PHPExcel();

        //Start adding next sheets
        $namesheet = array('List Data Kunjungan');
        //jumlah header
        $jumlah = array(1);

        $header = array(
            array(
                'Rekap Kunjungan Sales',
            )

        );
        $created_user = $this->session->userdata('user');
        // $area = $this->session->userdata('AREADISTRIBUTOR');
        //memisahkan multiple choice
        $tgl = explode("-", $this->input->post('tgl-sales'));
        $tgl1 = str_replace("/", "-", $tgl[0]);
        $tgl2 = str_replace("/", "-", $tgl[1]);
        $head_nama = $this->m_kunjungan->get_nama($tgl1, $tgl2, $created_user);
        $jumlahs = $this->m_kunjungan->get_kunjungan($tgl1, $tgl2, $created_user);
        $head_tgl = $this->m_kunjungan->get_header($tgl1, $tgl2, $created_user);
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


            $row = 2;
            $rownama = 3;
            $rowtrigger = 3;
            $rowjml = 3;

            if ($index == 0) {
                //grup nama_sales
                //    echo count($head_nama);
                //    die;
                foreach ($head_nama as $nama) {
                    $worksheet->setCellValue('A' . $rownama, $nama['nama_sales']);
                    $label[] = $nama['nama_sales'];
                    $rownama++;
                }
                //group by tanggal
                $worksheet->setCellValue('A' . $row, 'Nama');
                $char = 'B';


                foreach ($head_tgl as $h_tgl) {

                    $worksheet->setCellValue($char . $row, $h_tgl['tanggal']);
                    $rowjml = 3;
                    foreach ($head_nama as $trig) {
                        foreach ($jumlahs as $jml) {
                            if ($jml['tanggal'] == $h_tgl['tanggal']) {
                                if ($trig['nama_sales'] == $jml['nama_sales']) {
                                    $worksheet->setCellValue($char . $rowjml, $jml['jumlah']);
                                }
                            }
                        }
                        $rowjml++;
                    }



                    $char++;
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
        header('Content-Disposition: attachment;filename="Rekap Kunjungan.xlsx"');
        //unduh file
        $objWriter->save("php://output");
    }
  
}
