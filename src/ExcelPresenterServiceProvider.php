<?php

namespace Sands\Presenter\Excel;

use Illuminate\Support\ServiceProvider;

class ExcelPresenterServiceProvider extends ServiceProvider
{

    protected $presenters = [
        'xlsx' => [
            'presenter' => Presenters\Xlsx::class,
            'mimes' => [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ],
            'extensions' => [
                'xlsx',
            ],
        ],
        'xls' => [
            'presenter' => Presenters\Xls::class,
            'mimes' => [
                'application/vnd.ms-excel',
            ],
            'extensions' => [
                'xls',
            ],
        ],
        'csv' => [
            'presenter' => Presenters\Csv::class,
            'mimes' => [
                'text/csv',
            ],
            'extensions' => [
                'csv',
            ],
        ],
    ];

    /**
     * Bootstrap the application services.
     *
     * @return void
     */
    public function boot()
    {
        if (!$this->app->bound('excel')) {
            $this->app->register('Maatwebsite\Excel\ExcelServiceProvider');
        }
        if (!$this->app->bound('sands.presenter')) {
            $this->app->register('Sands\Presenter\PresenterServiceProvider');
        }
        $presenters = $this->presenters;
        $presenter = app('sands.presenter');
        foreach ($presenters as $name => $config) {
            $presenter->register($name, $config);
        }
        $presenter->setOption('excel.fileName', 'Export');
        $presenter->setOption('excel.title', '');
        $presenter->setOption('excel.description', '');
        $presenter->setOption('excel.creator', '');
        $presenter->setOption('excel.lastModifiedBy', '');
        $presenter->setOption('excel.subject', '');
        $presenter->setOption('excel.keywords', '');
        $presenter->setOption('excel.category', '');
        $presenter->setOption('excel.manager', '');
        $presenter->setOption('excel.company', '');
    }

    /**
     * Register the application services.
     *
     * @return void
     */
    public function register()
    {
    }
}
