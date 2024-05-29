<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;
use PhpOffice\PhpSpreadsheet\IOFactory;

class ExhibitorController extends Controller
{
    public function index()
    {
        $url = 'https://exhibitorlist.imsinoexpo.cn/front/index/index';
        $fromPageNumber = 0;
        $toPageNumber = 1;
        $recordsPerPage = 100;

        for ($page = $fromPageNumber; $page < $toPageNumber; $page++) {
            $data = [
                'area' => [],
                'cate_id' => [],
                'hall_id' => [],
                'keyword' => "",
                'lang' => 2,
                'page' => $page,
                'per_page' => $recordsPerPage,
                'session' => "",
                'uri' => "cphipmecchina2024shanghai",
            ];

            $response = Http::timeout(600)->withoutVerifying()->post($url, $data);

            if ($response->successful()) {
                // Request was successful, handle the response
                $responseData = $response->json();
                $this->addResultsToExcel($responseData);
            } else {
                // Request failed, handle the error
                $errorMessage = $response->body();
                dd($errorMessage);
            }
        }

        return 'Success! Pages from: ' . $fromPageNumber . ' - ' . $toPageNumber;
    }

    private function addResultsToExcel($response)
    {
        $resultsPath = public_path('results/exhibitorlist.xlsx');
        $spreadsheet = IOFactory::load($resultsPath);
        $sheet = $spreadsheet->getActiveSheet();
        $highestRow = $sheet->getHighestRow() + 1;

        $records = $response['data']['list']['data'];

        foreach ($records as $record) {
            $sheet->setCellValue('A' . $highestRow, $record['title_en']);
            $sheet->setCellValue('B' . $highestRow, $record['pos_no']);
            $sheet->setCellValue('C' . $highestRow, $record['area_en']);
            $sheet->setCellValue('D' . $highestRow, $record['cur_category_name_en']);

            // Increment row for the next record
            $highestRow++;
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($resultsPath);
    }
}
