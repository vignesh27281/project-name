<?php

namespace App\Http\Controllers;
use App\Exports\InvoiceExport;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;


class InvoiceController extends Controller
{
    public function export()
    {
        return Excel::download(new InvoiceExport, 'invoice.xlsx');
    }
}


