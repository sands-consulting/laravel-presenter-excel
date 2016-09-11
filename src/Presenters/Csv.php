<?php

namespace Sands\Presenter\Excel\Presenters;

use Sands\Presenter\Presenter;
use Sands\Presenter\PresenterContract;

class Csv implements PresenterContract
{
    public function __construct(Presenter $presenter)
    {
        $this->presenter = $presenter;
    }

    public function render($data = [])
    {
        $presenter = $this->presenter;
        return app('excel')->create($presenter->getOption('excel.fileName'), function ($excel) use ($presenter, $data) {
            $properties = [
                'title',
                'description',
                'creator',
                'lastModifiedBy',
                'subject',
                'keywords',
                'category',
                'manager',
                'company',
            ];
            foreach ($properties as $property) {
                $method = 'set' . ucfirst($property);
                $excel->$method($presenter->getOption("excel.{$property}"));
            }
            $excel->sheet('', function ($sheet) use ($data) {
                $sheet->fromArray($data, null, 'A1', true);
            });
        })->download('csv');
    }
}
