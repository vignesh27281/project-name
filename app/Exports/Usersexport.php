<?php

namespace App\Exports;

use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithHeadings;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class UsersExport implements FromCollection, WithStyles, WithColumnWidths
{
    public function collection()
{
    $specifications = [ 'COMMERCIAL TREADMILL'  ,' ',
       'Motor: 4 HP AC',    
       'Speed Range: 1.0-22 Km/hr',       
       'Walking Area: 1550 * 560 mm (61" * 22")',        
       'Dimension (L*W*H): 2120 * 900 *1590mm',        
       'Incline: 15 Level',        
       'Max User Weight: 150 Kgs',        
       'Display Type: White LED',        
       'Display Reading: Time,Speed,Distance,Calories, Pulse.',        
       'Program: 36 Preset program',        
       'Gross Weight: 190 Kgs',        
       'Features: Quick Speed&Incline button on Console',        
       'MP3 Can be played with wireless bluetooth Receiver.',        
       'Wheels for transportation easily',        
       '6-pieces elastic cushion system.',    
    ];
    $specifications2 = [ 'CURVE TREADMILL' , ' ',  
    'Running Area : 1600 * 580 mm',   
    'Display : LED',  
    'Program : User target' ,   
    'Reading :  Time,Speed,Calories,Distance,ODOMax' , 
    'User Weight : 190 kgAssembly Area (L*W*H): 1850 * 1050 * 650 mm' , 
    'Features: High stabillity,You can Easily move,High durability caterpillar belt.Anti corrosion Hardware.'
    ];

    $data = [        
       
        ['', '', '', 'QUOTATION'],
        ['To:', '', '', '', '', 'Quote Ref No : '],
        ['','Akshya Nagar 1st Block 1st Cross, Rammurthy nagar, Bangalore-560016', '', '', '', '', '2303CBE0057-Q1'],
        ['', '','', 'Proposed Products Specifications & Equipments '],
        [
            'S.NO.', 
            'Model', 
            'Image', 
            'Specifications', 
            'List/Unit Price',
             'Special Price', 
             'Qty', 
             'Total Amount'],
        [
          1,
         'WC-9900', 
         'Image.Png',
          implode("\n", $specifications),
          10000,
          5000,
          2, 
          10000],
        [ 
            2,
            'MP-8004',
            'Image2.jpg',
             implode("\n",  $specifications2),
             10000,
             5100,
             1,
             5100,],
        [''],
        [''],
        ['Terms & Conditions:'],
        ['','The prices are valid Only for 30 Days.'],
        ['Warranty:'],
        ['','* One Year Comprehensive Warranty for Parts and Labor.'],
        ['','* Warranty Doesn’t Cover Plastic, Rubber Parts, Upholstery and Physical Damages.'],
        ['','* Treadmill Warranty will be covered on use of Stabilizer.'],
        ['Delivery:'],
        ['','As per stock availability.'],
        ['Transportation:'],
        ['','* Transportation Charges Extra.'],
        ['','* Unloading Charges should be arranged by Customer.'],
        ['GST:'],
        ['','GST @ 18% included above'],
        ['Payment:'],
        ['','100% Advance Payment along with Purchase Order.'],
        ['Bank Details:'],
        ['','Bank Name : HDFC Bank'],
        ['','Account No : 50200043381896'],
        ['','Branch : TRICHY ROAD'],
        ['','IFSC code : HDFC0000031'],
        ['','GST NO. 33AAJCS4091D1ZH'],
        ['After Sales Support:'],
        ['','Dedicated Call center for service within 24 Hours of complaint registration'],
        [''],
        ['For S&T Welcare Equipments (P) Ltd:'],
        [''],
        ['Prakash','', '', '', '', 'Mahesh.C'],
        ['IT','', '', '', '', 'CEO/Director'],
        ['7094448225','','','','','','','']
    

    ];
    return collect($data);
}

    public function styles(Worksheet $sheet)
{
    return [
        'D1:P1' =>['font' =>['bold'=>true,'size'=>'18']],
        'A2:P4' => ['font' => ['bold' => true, 'size' => '13']],
        'A5:P5' => ['font' => ['bold' => true, 'size' => '10']],
        'A8:P8' => ['font' => ['size' => '10']],
        'A3:P3' => ['font' => ['bold' => true, 'size' => '10']],
        'B3:P9' => ['font' => ['size' => '10']],
        'C3' => ['font' => ['size' => '10']],
        'C3:P9' => ['font' => ['size' => '10']],
        'D4:P4' => ['bold' => true, 'font' => ['size' => '15']],
        'E3:P9' => ['font' => ['size' => '10']],
        'F3:P9' => ['font' => ['size' => '10']],
        'G3:P9' => ['font' => ['size' => '10']],
        'H3:P9' => ['font' => ['size' => '10']],
        'A1:P1' => ['font' => ['bold' => true, 'size' => '15'], 'alignment' => ['horizontal' => 'Center', 'vertical' => 'middle']],
        'A10:P10' => ['font' => ['bold' => true, 'size' => '13']],
        'A12:P12' => ['font' => ['bold' => true, 'size' => '13']],
        'A16:P16' => ['font' => ['bold' => true, 'size' => '13']],
        'A18:P18' => ['font' => ['bold' => true, 'size' => '13']],
        'A21:P21' => ['font' => ['bold' => true, 'size' => '13']],
        'A23:P23' => ['font' => ['bold' => true, 'size' => '13']],
        'A25:P25' =>['font'=> ['bold' =>true, 'size'=>'13']],
        'A31:P31' =>['font'=> ['bold' =>true, 'size'=>'13']],
        'A34:P34' =>['font'=> ['bold' =>true, 'size'=>'13']],
        'A36:P36' =>['font'=> ['bold' =>true, 'size'=>'13']],
        'A37:P37' =>['font'=> ['bold' =>true, 'size'=>'13']],
        'A38' =>['font'=> ['bold' =>true, 'size'=>'13'], 'alignment' => ['horizontal' => 'left']],
    ];
 
}


public function columnWidths(): array
{
    return [
        
        'A' => 7.26,
        'B' => 12.73,
        'C' => 37.82,
        'D' => 76.36,
        'E' => 13.40,
        'F' => 15.56,
        'G' => 8.91,
        'H' => 15.40,
       
    ];
}
   
}

