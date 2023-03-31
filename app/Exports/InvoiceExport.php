<?php

namespace App\Exports;

use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithHeadings;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class InvoiceExport implements FromCollection, WithStyles, WithColumnWidths
{
    public function collection()
{
    
    $data = [        
       
        ['S AND T WELCARE EQUIPMENTS PRIVATE LIMITED','','','','','OUR GST.NO: 33AAJCS4091D1ZH'],
        ['MADURAI OFFICE:59,SIVABHAYA COMPLEX                '],
        ['OPPOSITE TO THEEKATHIR PRESS ,VILANGUDI MADURAI-625018        '],
        ['PHONE NUMBER: 9965555555              '],
        ['Name','Prakash','','','xxxx','','Date:','27.03.2023'],
        ['Address','Peelamadu','','ORDER NO:','xxxx','','Date:','27.03.2023'],
        ['City','Coimbatore','','MOBILE NO;','','xxxx'],
        ['State','TN','','OTHER CONTACT NO:','','xxxx'],
        ['Pincode','641004','','DEL DATE','','xxxx'],
        ['LandMark','','','DEL FROM','','xxxx'],
        ['Branch Name','MADURAI','','EXECUTIVE NAME','','xxxx'],
        ['GST NO : ','','','EMPLOYEE CODE','','xxxx'],
        ['HNA COMMISSION APPROVAL CODE'],
        ['PAYMENT DELAY APPROVAL CODE'],
        ['PERSONAL INFORMATION'],
        ['USER :'],
        ['AGE :','','','NO. OF MEMBERS :'],
        ['WEIGHT :','','','OCCUPATION :'],
        ['HOW THE CUSTOMER KNOW ABOUT WELCARE : Extn Cust','','','PURPOSE OF BUYING : ','Weightloss'],
        [''],
        ['Sl No ','Model No','Description','','Least price','Qty','Selling price','Amount'],
        [1,'W152','SCALE','','1','1000.00','847.46',''],
        [2,'2.5KG','ROUND DUMBULLES','','2','900.00','762.61'],
        ['','','','','','','0.00'],
        [''],[''],[''],[''], [''], [''],[''],[''],[''],[''],[''],[''], [''],[''],
        ['','','Gross Total','','','','','1610.17'],
        ['','','GST @ 18%','','','','','289.83'],
        ['','','Round off(+/-)','','','','','0.00'],
        ['','','Nett Total  ','','','','','1900.00'],
        ['','','Advance Amount:  ','','CARD','','','900.00'],
        ['Rupees:','','TWO THOUSAND SIX HUNDRED  ','','','','',''],
        ['','','Balance Total','','','','','1000'],
        ['1.Central sales tax of general sales tax will be charged at the applicable rate ','','','','','','','S & T WELCARE EQUIPMENTS '],
        ['if your declaration forms are not accepted by the sales tax Authorities at the time assessment.','','  ','','','','','P LTD'],
        ['2.Goods once sent according to order will not be taken back'],
        ['3.All disputes are subject to Bangalore Jurisdiction'],
        ['E & O E','','  ','','','','','Authorised Signatory'],

    

    ];

    return collect($data);
}

    public function styles(Worksheet $sheet)
{
    return [
        'A1:P1' => ['font' => ['bold' => true, 'size' => '13'], 'alignment' => ['horizontal' => 'left', 'vertical' => 'middle' ]], 
         
    ];
 
}


public function columnWidths(): array
{
    return [
        
        'A' => 16.43,
        'B' => 15.80,
        'C' => 15.55,
        'D' => 19.40,
        'E' => 17.54,
        'F' => 3.75,
        'G' => 13.19,
        'H' => 10.58,
       
    ];
}

   
}


