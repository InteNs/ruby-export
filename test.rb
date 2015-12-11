gem 'axlsx', '~> 2.0.1'

require 'axlsx'
require 'ostruct'
require 'pp'

def headers
  [
  :wagon_nr,
  :load_holder_nr,
  :cargo_nr,
  :dangerous_good_nr
  ]
end

data = [
  {wagon_nr: 'foo',
   load_holder: [
     {load_holder_nr: 'bar',
      cargos: [
        {cargo_nr: 'baz',
         dangerous_goods: [
           {dangerous_good_nr: 'baq'
        }]
     }]
  }]
}]

def build_rows(attributes)
  main_row = Array.new(headers.count)
  attributes.each_with_object([]) do |(attribute, value), rows|
    if value.is_a?(Array)
      nested_rows = build_nested_rows(value)
      main_row = static_merge(main_row, nested_rows.first)
    else
      add_to_row(main_row, attribute, value)
    end
  end.unshift(main_row)
end

def add_to_row(row, attribute, value)
  row[index_of(attribute) || row.count + 1] = value
end

def build_matrix(data)
  data.each_with_object([]) do |attributes, rows|
    build_rows(attributes).each { |row| rows << row }
  end
end

def build_nested_rows(value)
  value.map { |attributes| build_rows(attributes) }.first
end

def static_merge(array1, array2)
  array1.zip(array2).map { |xs| xs.compact.first }
end

def index_of(attribute)
  headers.find_index(attribute)
end

def first_iteration?(iteration)
  iteration == 0
end

Axlsx::Package.new do |p|
  p.workbook.add_worksheet(name: 'Pie Chart') do |sheet|
    sheet.add_row headers
    build_matrix(data).each do |row_data|
      sheet.add_row(row_data)
    end
    # sheet.add_table "A1:C4", :name => 'Build Matrix', :style_info => { :name => "TableStyleMedium23" }
  end
  p.serialize 'testresult.xlsx'
end
