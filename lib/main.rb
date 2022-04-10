# frozen_string_literal: true

require 'axlsx'
require 'fast_excel'
require 'json'
require 'benchmark'
require 'warning'
Warning.ignore(/deprecated/)

class ExcelExport
  HEADERS = [
    'country',
    'name',
    'latitude',
    'longitude'
  ]

  def initialize
    @data = JSON.parse File.read("#{File.dirname(__FILE__)}/cities.json")
  end

  def fast_excel
    workbook = FastExcel.open('export_fast_excel.xlsx')
    worksheet = workbook.add_worksheet
    worksheet.append_row(HEADERS, )
    @data.each do |city|
      worksheet.append_row(
        [
          city['country'],
          city['name'],
          city['latitude'],
          city['longitude']
        ]
      )
    end
    workbook.close
  end

  def axlsx
    p = Axlsx::Package.new
    p.workbook.add_worksheet(name: 'Basic Worksheet') do |sheet|
      sheet.add_row HEADERS
      @data.each do |city|
        sheet.add_row(
          [
            city['country'],
            city['name'],
            city['latitude'],
            city['longitude']
          ]
        )
      end
    end
    p.serialize('export_axlsx.xlsx')
  end
end

exporter = ExcelExport.new
Benchmark.bm do |benchmark|
  benchmark.report('AX: ') { exporter.axlsx }
  benchmark.report('FE: ') { exporter.fast_excel }
end
