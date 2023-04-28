<?php
defined('BASEPATH') or exit('No direct script access allowed');

class C_rekap_sales extends CI_Controller
{
    public function __construct()
    {
        parent::__construct();
        $this->ci = &get_instance();
        $this->load->library('session');
        $this->load->library(array('PHPExcel_/PHPExcel', 'PHPExcel_/PHPExcel/IOFactory'));
        $this->load->helper('url');
        $this->load->helper('form');
        $this->load->model('m_sales');



        ini_set('max_execution_time', 0);
        ini_set('memory_limit', '2048M');
    }
    public function index()
    {
        //membuat objek PHPExcel
        $objPHPExcel = new PHPExcel();

        //Start adding next sheets
        $namesheet = array('List Data Seles');
        //jumlah header
        $jumlah = array(4);

        $header = array(
            array(
                'Nama Karyawan',
                'Target Bulanan',
                'Target Harian',
                'Target Sampai Hari',
            )

        );
        $created_user = $this->session->userdata('user');
        // $area = $this->session->userdata('AREADISTRIBUTOR');
        //memisahkan multiple choice
        $tgl = explode("-", $this->input->post('tgl-sales'));
        $tgl1 = str_replace("/", "-", $tgl[0]);
        $tgl2 = str_replace("/", "-", $tgl[1]);
        $namas = $this->m_sales->get_nama();
        $subnamas = $this->m_sales->get_subnama();
        $targets = $this->m_sales->get_target();
        $totals = $this->m_sales->get_total();
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

            if ($index == 0) {
                $rowhead = 2;
                // $char = 'C';
                foreach ($namas as $nama) {
                    $worksheet->setCellValue('A' . $rowhead, $nama['nama_leader']);
                    $worksheet->getStyle('A' . $rowhead)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $worksheet->getStyle('A' . $rowhead)->getFill()->getStartColor()->setARGB('808080');
                    $worksheet->setCellValue('B' . $rowhead, '1.111.111');
                    $worksheet->setCellValue('C' . $rowhead, '555.555');
                    foreach ($totals as $total) {
                        if ($total['id_leader'] == $nama['id_leader']) {
                            $worksheet->setCellValue('D' . $rowhead, $total['total']);
                        }
                    }
                    $rowhead++;
                    // $char = 'C';
                    foreach ($subnamas as $subnama) {
                        if ($subnama['id_leader'] == $nama['id_leader']) {
                            $worksheet->setCellValue('A' . $rowhead, $subnama['nama_sales']);
                            $worksheet->setCellValue('B' . $rowhead, '1.000.000');
                            $worksheet->setCellValue('C' . $rowhead, '500.000');
                            foreach ($targets as $target) {
                                if ($subnama['id_sales'] == $target['id_sales'])
                                    $worksheet->setCellValue('D' . $rowhead, $target['total']);
                            }
                            $rowhead++;
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
        header('Content-Disposition: attachment;filename="Rekap Transaksi Sales.xlsx"');
        //unduh file
        $objWriter->save("php://output");
    }
}
