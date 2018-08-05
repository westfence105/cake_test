<?php
namespace App\Controller;

use App\Controller\AppController;

use Cake\Event\Event;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as XlsxWriter;
use PhpOffice\PhpSpreadsheet\Style as SheetStyle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Office Controller
 *
 *
 * @method \App\Model\Entity\Office[]|\Cake\Datasource\ResultSetInterface paginate($object = null, array $settings = [])
 */
class OfficeController extends AppController
{
    public function beforeFilter( Event $ev ){
        $this->loadModel('Users');
        $this->loadModel('Companies');
        $this->loadModel('Colors');

        $this->response->type([
            'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ]);
    }

    public function index() {

    }

    public function excel01()
    {

        $spreadsheet = new Spreadsheet();

        $spreadsheet->getProperties()
                    ->setTitle('Test')
                    ->setSubject('SpreadSheet')
                    ->setCreator('Togo Nishigaki')
                    ->setCompany('soundbag')
                    ->setManager('Admin')
                    ->setCategory('test_data')
                    ->setDescription('PhpSpreadsheet Test')
                    ->setKeywords('test')
                ;

        {
            $users_query = $this->Users->find('all');
            $data_01 = [
                ['#', '氏名','性別','生年月日','郵便番号','都道府県','電話番号','EMail']
            ];

            foreach( $users_query as $user ){
                $data_01[] = [
                    $user->id,
                    $user->name,
                    $user->gender,
                    $user->birth->format('Y-m-d'),
                    $user->postalcode,
                    $user->prefecture,
                    $user->tel,
                    $user->email,
                ];
            }

            $row_max = $users_query->count() + 1;

            $users_sheet = $spreadsheet->getSheet(0);

            $users_sheet->setTitle('Users');
            $users_sheet->fromArray( $data_01, NULL, 'A1' );

            $users_sheet->getStyle('A1:A'.$row_max)->getAlignment()->setHorizontal( SheetStyle\Alignment::HORIZONTAL_RIGHT );
            $users_sheet->getStyle('B1:F1')->getAlignment()->setHorizontal( SheetStyle\Alignment::HORIZONTAL_CENTER );
            $users_sheet->getStyle('C1:C'.$row_max)->getAlignment()->setHorizontal( SheetStyle\Alignment::HORIZONTAL_CENTER );
            $users_sheet->getStyle('F1:F'.$row_max)->getAlignment()->setHorizontal( SheetStyle\Alignment::HORIZONTAL_CENTER );

            $users_sheet->getColumnDimension('B')->setWidth(14);
            $users_sheet->getColumnDimension('C')->setWidth( 5);
            $users_sheet->getColumnDimension('D')->setWidth(14);
            $users_sheet->getColumnDimension('E')->setWidth(14);
            $users_sheet->getColumnDimension('F')->setWidth(12);
            $users_sheet->getColumnDimension('G')->setWidth(20);
            $users_sheet->getColumnDimension('H')->setWidth(28);
        }

        {
            $companies_query = $this->Companies->find('all');
            $data_02 = [
                ['#', '企業名','住所','電話番号','EMail']
            ];

            foreach( $companies_query as $company ){
                $data_02[] = [
                    $company->id,
                    $company->name,
                    $company->address,
                    $company->phone,
                    $company->email,
                ];
            }

            $row_max = $companies_query->count() + 1;

            $companies_sheet = new Worksheet( $spreadsheet, 'Companies' );
            $spreadsheet->addSheet( $companies_sheet );
            $companies_sheet->fromArray( $data_02, NULL, 'A1' );

            $companies_sheet->getStyle('A1:A'.$row_max)->getAlignment()->setHorizontal( SheetStyle\Alignment::HORIZONTAL_RIGHT );

            $companies_sheet->getColumnDimension('B')->setWidth(18);
            $companies_sheet->getColumnDimension('C')->setWidth(24);
            $companies_sheet->getColumnDimension('D')->setWidth(24);
            $companies_sheet->getColumnDimension('E')->setWidth(36);
        }

        {
            $colors_sheet = new Worksheet( $spreadsheet, 'Colors' );
            $spreadsheet->addSheet( $colors_sheet );

            $colors_query = $this->Colors->find('all');

            $i = 1;
            foreach( $colors_query as $color ){
                $colors_sheet->setCellValue('A'.$i, $color->name );
                $colors_sheet->setCellValue('B'.$i, $color->code );
                $colors_sheet->getStyle('B'.$i )->getFill()->setFillType( SheetStyle\Fill::FILL_SOLID );
                $colors_sheet->getStyle('B'.$i )->getFill()->getStartColor()->setRGB( $color->code );
                if( $color->name == 'black' ){
                    $colors_sheet->getStyle('B'.$i )->getFont()->getColor()->setRGB('ffffff');
                }
                ++$i;
            }
        }

        {
            $fibonacci_sheet = new Worksheet( $spreadsheet, 'fibonacci' );
            $spreadsheet->addSheet( $fibonacci_sheet );

            $fibonacci_sheet->setCellValue('A1', 0 );
            $fibonacci_sheet->setCellValue('A2', 1 );
            for ( $i = 3;  $i < 25;  $i++) {
                $fibonacci_sheet->setCellValue('A'.$i, '=SUM(A'.( $i - 2 ).':A'.( $i - 1 ).')' );
            }
        }

        ob_start();
        $writer = new XlsxWriter( $spreadsheet );
        $writer->save('php://output');

        $response = $this->response;
        $response->body( ob_get_contents() );
        $response = $response->withType('xlsx')
                             ->withDownload('excel_01.xlsx')
                           ;

        ob_clean();
        return $response;
    }
}
