<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExcelController extends Controller
{  

    # Mảng Cột số liệu
    const arr_colum_so_lieu = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];

    # Mảng Cột số liệu
    const arr_colum_so_lieu_thuc_te = ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA'];

    # Mảng Cột tổng hơp
    const arr_colum_tong_hop = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];

    # Mảng Cột Công Nghiệp
    const arr_colum_so_lieu_cn = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];

    # Mảng Cột Xây dựng
    const arr_colum_so_lieu_xay_dung = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];

    # Mảng Cột Xây dựng
    const arr_colum_so_lieu_dich_vu = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
    
    # Mảng Cột cây trồng
    const arr_colum_so_lieu_cay_trong = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
    
    # Một tỷ
    const mot_ty = 1000000000;
    
    # Mảng index số liệu
    private $array_index_So_Lieu = array(
                # Cây trồng
                array(
                    'start' => 7,
                    'end'   => 31
                ),
                # Chăn nuôi
                array(
                    'start' => 36,
                    'end'   => 50
                ),
                # Dịch vụ trồng trọt
                array(
                    'start' => 53,
                    'end'   => 60
                ),
                # Dịch vụ chăn nuôi
                array(
                    'start' => 62,
                    'end'   => 67
                ),
                # Trồng và nuôi rừng
                array(
                    'start' => 70,
                    'end'   => 75
                ),
                # Khai thác
                array(
                    'start' => 77,
                    'end'   => 85
                ),
                # Thu nhặc từ rừng
                array(
                    'start' => 87,
                    'end'   => 91
                ),
                # Hoạt động lâm nghiệp
                array(
                    'start' => 93,
                    'end'   => 95
                ),            
                # Thủy sản
                array(
                    'start' => 97,
                    'end'   => 100
                ),
    );

    # Mảng index tổng hợp
    private $array_index_Tong_Hop = array(
        # Cây trồng
        array(
            'start' => 9,
            'end'   => 33
        ),
        # Chăn nuôi
        array(
            'start' => 35,
            'end'   => 49
        ),
        # Dịch vụ trồng trọt
        array(
            'start' => 52,
            'end'   => 59
        ),
        # Dịch vụ chăn nuôi
        array(
            'start' => 61,
            'end'   => 66
        ),
        # Trồng và nuôi rừng
        array(
            'start' => 69,
            'end'   => 74
        ),
        # Khai thác
        array(
            'start' => 76,
            'end'   => 84
        ),
        # Thu nhặc từ rừng
        array(
            'start' => 86,
            'end'   => 90
        ),
        # Hoạt động lâm nghiệp
        array(
            'start' => 92,
            'end'   => 94
        ),            
        # Thủy sản
        array(
            'start' => 96,
            'end'   => 99
        ),
    );

    # Mảng index tổng hợp thực tế
    private $array_index_Tong_Hop_Thuc_Te = array(
        # Cây trồng
        array(
            'start' => 108,
            'end'   => 132
        ),
        # Chăn nuôi
        array(
            'start' => 134,
            'end'   => 148
        ),
        # Dịch vụ trồng trọt
        array(
            'start' => 151,
            'end'   => 158
        ),
        # Dịch vụ chăn nuôi
        array(
            'start' => 160,
            'end'   => 165
        ),
        # Trồng và nuôi rừng
        array(
            'start' => 168,
            'end'   => 173
        ),
        # Khai thác
        array(
            'start' => 175,
            'end'   => 183
        ),
        # Thu nhặc từ rừng
        array(
            'start' => 185,
            'end'   => 189
        ),
        # Hoạt động lâm nghiệp
        array(
            'start' => 191,
            'end'   => 193
        ),            
        # Thủy sản
        array(
            'start' => 195,
            'end'   => 198
        ),
    );

    # Mảng số liệu cho công nghiệp xây dựng
    private $array_index_So_Lieu_CNXD = array(
        # Công nghiệp
        array(
            'start' => 19,
            'end'   => 21
        ),
    );

    public function XuLyDuLieu(){
        try {
            $url_report = 'reports/PHONG.SO_LIEU_15-25_(ĐA_THAM_DINH)-0782019_(huyen)-CSG_moi.xls';
            $spreadsheet = IOFactory::load($url_report);

            #### Lọc giá trị lưu vô mảng 2 chiều từ sheet TINH GTSX NLT
            $sheet = $spreadsheet->getSheetByName('TINH GTSX NLT ');
            $sheet_cay_trong = $spreadsheet->getSheetByName(' CAY TRONG');

            ##### Đọc dữ liệu lưu lại
            $url_tong_hop_gia_tri = 'reports/TONG_HOP_GIA_TRI_KTXH_15-25.xlsx';
            $spreadsheet_TH = IOFactory::load($url_tong_hop_gia_tri);
            $sheet_tong_hop_gia_tri = $spreadsheet_TH->getActiveSheet();

            #### 
            $sheet_cong_nghiep = $spreadsheet->getSheetByName('CN');
            $sheet_xay_dung = $spreadsheet->getSheetByName('XD');
            $sheet_dich_vu = $spreadsheet->getSheetByName('DV');
            $sheet_tong_hop_gia_tri = $spreadsheet_TH->getActiveSheet();
            
            ## Xử lý N-L-T
            $array = array();
            $this->XuLySoLieuNongLamThuy($sheet, $sheet_tong_hop_gia_tri, $array);

            ## XỬ lý CNXD DV
            $this->XuLySoLieuCNXD_DV($sheet_cong_nghiep, $sheet_xay_dung, $sheet_dich_vu, $sheet_tong_hop_gia_tri);
        
            ## Render diện tích
            $this->XuLyTongDienTich($sheet_cay_trong, $sheet_tong_hop_gia_tri);

            # Viết ra file
            $writer = new Xlsx($spreadsheet_TH);
            $writer->save($url_tong_hop_gia_tri);
            //throw new Exception('Lỗi nè');
        } 
        catch (Exception $e) {
            return response()->json(['data'=>'Có lỗi xảy ra', 'error'=>$e]);
        }
        
        return response()->json(['data'=>'Thành công']);
    }

    ## Xử lý số liệu Công nghiệp xây dựng & dịch vụ
    public function XuLySoLieuCNXD_DV($spreadsheet_cong_nghiep, $spreadsheet_xay_dung, $spreadsheet_dich_vu, $spreadsheet_tong_hop_gia_tri){
        $row_cn_2010         = 101;
        $row_xay_dung_2010   = 11;
        $row_tong_hop_dich_vu_2010   = 103;
        $row_tong_hop_xay_dung_2010   = 102;
        $row_tong_hop_cn_thuc_te         = 200;
        $row_xay_dung_thuc_te   = 6;
        $row_tong_hop_dich_vu_thuc_te   = 202;
        $row_tong_hop_xay_dung_thuc_te   = 201;
        for($i = 0; $i< count(self::arr_colum_so_lieu_cn); $i++){
            $sum_2010 = 0;
            $sum_thuc_te = 0;
            for($j = $this->array_index_So_Lieu_CNXD[0]['start']; $j <=  $this->array_index_So_Lieu_CNXD[0]['end']; $j++){
                $sum_2010 += $spreadsheet_cong_nghiep->getCell(self::arr_colum_so_lieu_cn[$i].$j)->getCalculatedValue();
            }
            for($j =9; $j <= 11; $j++){
                $sum_thuc_te += $spreadsheet_cong_nghiep->getCell(self::arr_colum_so_lieu_cn[$i].$j)->getCalculatedValue();
            }
            ## Công nghiệp xây dựng
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_cn_2010,round($sum_2010/1000 , 2, PHP_ROUND_HALF_UP)); 
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_cn_thuc_te,round($sum_thuc_te/1000 , 2, PHP_ROUND_HALF_UP)); 

            ## Xây dựng
            $gia_tri_xay_dung = $spreadsheet_xay_dung->getCell(self::arr_colum_so_lieu_xay_dung[$i].$row_xay_dung_2010)->getCalculatedValue();
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_xay_dung_2010, $gia_tri_xay_dung); 
            
            $gia_tri_xay_dung_thuc_te = $spreadsheet_xay_dung->getCell(self::arr_colum_so_lieu_xay_dung[$i].$row_xay_dung_thuc_te)->getCalculatedValue();
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_xay_dung_thuc_te, $gia_tri_xay_dung_thuc_te); 
            #echo($gia_tri_xay_dung_thuc_te.'<br>');

            $sum_2010 = 0;
            $sum_thuc_te = 0;
            for($j = 35; $j<= 49; $j++){
              $sum_2010 +=   $spreadsheet_dich_vu->getCell(self::arr_colum_so_lieu_dich_vu[$i].$j)->getCalculatedValue();
            }

            for($j = 5; $j<= 19; $j++){
                $sum_thuc_te +=   $spreadsheet_dich_vu->getCell(self::arr_colum_so_lieu_dich_vu[$i].$j)->getCalculatedValue();
            }

            ## Dịch vụ
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_dich_vu_2010, $sum_2010); 
            $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_dich_vu_thuc_te, $sum_thuc_te); 
            $sum_2010 = 0;
            $sum_thuc_te = 0;

        }
    }

    ## Xử lý số liệu Nông lâm thủy sản
    public function XuLySoLieuNongLamThuy($sheet_nong_lam_thuy, $spreadsheet_tong_hop_gia_tri, $array){

        for($i = 0; $i< count($this->array_index_So_Lieu); $i++){
            for($j = $this->array_index_So_Lieu[$i]['start']; $j <= $this->array_index_So_Lieu[$i]['end']; $j++){
                # Khởi tạo mảng
                $child_array = array();
                # Lấy giá SS 2010 từ ô O[$i]
                $giass_2010         = $sheet_nong_lam_thuy->getCell('O'.$j)->getCalculatedValue();
                for($k = 0; $k < count(self::arr_colum_so_lieu); $k++){
                    # Lấy data
                    $value = ($sheet_nong_lam_thuy->getCell(self::arr_colum_so_lieu[$k].$j)->getCalculatedValue()*$giass_2010)/self::mot_ty;
                    # Làm tròn 2 chữ số
                    $Round = round($value , 2, PHP_ROUND_HALF_UP);
                    # Đẩy giá trị vào mảng từ dòng [$J][$i] trong file excel
                    array_push($child_array,$Round);
                }
                # Đẩy vô mảng 2 chiều
                array_push($array, $child_array);   
            }
        }


        for($i = 0; $i< count($this->array_index_Tong_Hop); $i++){
            for($j = $this->array_index_Tong_Hop[$i]['start']; $j <= $this->array_index_Tong_Hop[$i]['end']; $j++){
                for($k = 0; $k < count(self::arr_colum_tong_hop); $k++){
                    $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$k].$j, $array[0][$k]);
                }
                $remove = array_shift($array);  
            }
        }

        $array = array();

        for($i = 0; $i< count($this->array_index_So_Lieu); $i++){
            //echo($this->array_index_So_Lieu[$i]['start'].'=>'.$this->array_index_So_Lieu[$i]['end'].'<br>');
            for($j = $this->array_index_So_Lieu[$i]['start']; $j <= $this->array_index_So_Lieu[$i]['end']; $j++){
                # Khởi tạo mảng
                $child_array = array();
                # Lấy giá SS 2010 từ ô O[$i]
                $giass_2010         = $sheet_nong_lam_thuy->getCell('O'.$j)->getCalculatedValue();
                for($k = 0; $k < count(self::arr_colum_so_lieu); $k++){
                    # Lấy data
                    $value = ($sheet_nong_lam_thuy->getCell(self::arr_colum_so_lieu[$k].$j)->getCalculatedValue() * $sheet_nong_lam_thuy->getCell(self::arr_colum_so_lieu_thuc_te[$k].$j)->getCalculatedValue() * $giass_2010)/self::mot_ty;
                    # Làm tròn 2 chữ số
                    $Round = round($value , 2, PHP_ROUND_HALF_UP);
                    # Đẩy giá trị vào mảng từ dòng [$J][$i] trong file excel
                    array_push($child_array,$Round);
                }
                # Đẩy vô mảng 2 chiều
                array_push($array, $child_array);   
            }
        }

        for($i = 0; $i< count($this->array_index_Tong_Hop_Thuc_Te); $i++){
            for($j = $this->array_index_Tong_Hop_Thuc_Te[$i]['start']; $j <= $this->array_index_Tong_Hop_Thuc_Te[$i]['end']; $j++){
                for($k = 0; $k < count(self::arr_colum_tong_hop); $k++){
                    $spreadsheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$k].$j, $array[0][$k]);
                }
                $remove = array_shift($array);  
            }
        }
    }

    public function XuLyTongDienTich($sheet_cay_trong, $sheet_tong_hop_gia_tri){
        $row_so_lieu_tong_dien_tich = 10;
        $row_tong_hop_tong_dien_tich = 213;
        for($i = 0; $i < count(self::arr_colum_so_lieu_cay_trong); $i++){
            $value = $sheet_cay_trong->getCell(self::arr_colum_so_lieu_cay_trong[$i].$row_so_lieu_tong_dien_tich)->getCalculatedValue();
            $sheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_tong_dien_tich, $value);
        }
    }

    public function test(){
        // $url_report = 'reports/PHONG.SO_LIEU_15-25_(ĐA_THAM_DINH)-0782019_(huyen)-CSG_moi.xls';
        // $spreadsheet = IOFactory::load($url_report);
        // $array = array(37, 57, 77, 81, 97, 114, 118, 123, 132, 156, 161, 168, 180, 187, 195, 222, 232,
        // 239, 246, 254, 260, 266);
        // $sheet_cay_trong = $spreadsheet->getSheetByName(' CAY TRONG');

        // ##### Đọc dữ liệu lưu lại
        // $url_tong_hop_gia_tri = 'reports/TONG_HOP_GIA_TRI_KTXH_15-25.xlsx';
        // $spreadsheet_TH = IOFactory::load($url_tong_hop_gia_tri);
        // $sheet_tong_hop_gia_tri = $spreadsheet_TH->getActiveSheet();

        // $row_so_lieu_tong_dien_tich = 10;
        // $row_tong_hop_tong_dien_tich = 213;
        // for($i = 0; $i < count(self::arr_colum_so_lieu_cay_trong); $i++){
        //     $value = $sheet_cay_trong->getCell(self::arr_colum_so_lieu_cay_trong[$i].$row_so_lieu_tong_dien_tich)->getCalculatedValue();
        //     $sheet_tong_hop_gia_tri->setCellValue(self::arr_colum_tong_hop[$i].$row_tong_hop_tong_dien_tich, $value);
        // }

        // # Viết ra file
        // $writer = new Xlsx($spreadsheet_TH);
        // $writer->save($url_tong_hop_gia_tri);
    }
}
