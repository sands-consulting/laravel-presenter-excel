# Sands\Presenter\Excel

[Sands Laravel Presenter](https://github.com/sands-consulting/laravel-presenter) plugin to respond with MS Excel files: XLSX, XLS and CSV. Files are generated using the awesome [Laravel Excel](https://github.com/Maatwebsite/Laravel-Excel) package.

## Installation

```bash
$ composer require sands/laravel-presenter-excel
```
Make sure [Sands\Presenter](https://github.com/sands-consulting/laravel-presenter) has been properly installed.

In `config/app.php` add `Sands\Presenter\Excel\ExcelPresenterServiceProvider` after the `Sands\Presenter\PresenterServiceProvider` inside the `providers` array:

```php
'providers' => [
     ...
     Sands\Presenter\PresenterServiceProvider::class,
     Sands\Presenter\Excel\ExcelPresenterServiceProvider::class,
     ...
]
```

## Usage

This plugin allows you to easily create XLSX, XLS and CSV exports. 

Let's say you have a UsersController with the below method:

```php
public function index()
{
    return $this->present(['users' => User::all()])
        ->setOption('excel.fileName', 'User Reports')
        ->using('blade', 'xlsx', 'csv', 'xls');
}
```
When the user visits `/users?format=xlsx`, they will be prompted to download the `User Reports.xlsx` file.

When the user visits `/users?format=xls`, they will be prompted to download the `User Reports.xls` file.

When the user visits `/users?format=csv`, they will be prompted to download the `User Reports.csv` file.

### Data Format for XLSX and XLS

Data for XLSX and XLS must be a nested collection as per example below:

```php
public function csvData()
{
    return [
        "Sheet 1 Name" => [            
            [row 1 data],
            [row 2 data],
            ...
        ],
        ...
    ];
    
    ... or ...
    
    return [
        "Users" => User::all(),
        "Tasks" => Task::all(),
        ...
    ];
}
```

### Data Format for CSV

Data for CSV must be as collection per example below:

```php
public function csvData()
{
    return [
        [row 1 data],
        [row 2 data],
        ...
    ];
    
    ... or ...
    
    return User::all();
}
```

### File Properties

File properties can be set using the `setOption` method. Properties available are as per [defined here](http://www.maatwebsite.nl/laravel-excel/docs/reference-guide).

The export file name can be changed via the `excel.fileName` option.

```php
public function index()
{
    return $this->present()
        ->setOption('data.blade', 'bladeData')
        ->setOption('data.xlsx', 'data')
        ->setOption('excel.fileName', 'User Reports')
        ->setOption('excel.title', 'Exported User Reports')
        ->using('blade', 'xlsx');
}
```

## MIT License

Copyright (c) 2016 Sands Consulting Sdn Bhd


Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.