gem 'axlsx', '~> 2.0.1'

require 'axlsx'
require 'ostruct'

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

def build_row(data, headers)
  data.first.each_with_object([]) do |(attribute, value), row|
    index = headers.find_index(attribute) if headers.include?(attribute)
end
Axlsx::Package.new do |p|
  p.workbook.add_worksheet(name: 'Pie Chart') do |sheet|
    sheet.add_row headers
    sheet.add_row build_row(data, headers)
    # sheet.add_table "A1:C4", :name => 'Build Matrix', :style_info => { :name => "TableStyleMedium23" }
  end
  p.serialize 'testresult.xlsx'
end

def flatten_hash(hash)
  hash.each_with_object([]) do |(key, value), result|
    if value.is_a? Array
      # do stuff
    else
      (result.first ||= {})
