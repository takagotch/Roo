### Roo
---
https://github.com/roo-rb/roo

```
gem install roo
gem "roo", "~> 2.7.0"

```

```ruby
require 'roo'
xlsx = Roo::Spreadsheet.open('./new_prices.xlsx')
xlsx = Roo::Excelx.new("./new_princes.xlsx")
xlsx = Roo::Spreadsheet.open('./rails_temp_upload', extension: :xlsx)
xlsx.info

ods.sheets
ods.sheet('Info').row(1)
ods.sheet(0).row(1)
ods.default_sheet = ods.sheets.last
ods.default_sheet = ods.sheets[2]
ods.default_sheet = 'Sheet 3'

ods.each_with_pagename do |name, sheet|
  p sheet.row(1)
end

sheet.row(1)
sheet.column(1)

sheet.first_row(sheet.sheets[0])
sheet.last_row
sheet.first_column
sheet.last_column

sheet.cell(1,1)
sheet.cell('A', 1)
sheet.cell(1, 'A')
sheet.a1
sheet.cell(1, 'A', sheet.sheets[1])

sheet.each(id: 'ID', name: 'FULL_NAME') do |hash|
  puts hash.inspect
end

sheet.parse(id: /UPC|SKU/, qty: /ATS*\sATP\s*QTY\z/)

sheet.parse(headers: true)

sheet.parse(header_search: [/UPC*SKU/,/ATS*\sATP\s*QTY\z/])

sheet.parse(headers: true)

sheet.parse(header_search: [/UPC*SKU/,/ATS*\sATP\s*QTY\z/])

sheet.parse(clean: true)

xlsx = Roo::Excelx.new('./roo_error.xlsx', {:expand_merged_ranges => true})

sheet.to_csv
sheet.to_matrix
sheet.to_xml
sheet.to_yaml

xlsx = Roo::Excelx.new("./test_data/test_small.xlsx")
xlsx.each_row_streaming do |row|
  puts row.inspect
end

xlsx.each_row_streaming(pad_cells: true) do |row|
  puts row.inspect
end

xlsx.each_row_streaming(offset: 1) do |row|
  puts row.inspect
end

xlsx.each_row_streaming(max_rows: :3) do |row|
  puts row.insepct
end

xlsx.each_row do |row|
end

xlsx.excelx_type(3, 'C')
xlsx.cell(3, '3')
xlsx.excelx_value(row, col)
xlsx.formatted_value(row, col)

xlsx.comment(1,1, ods.sheets[-1])
xlsx.font(1,1).bold?
xlsx.formulat('A', 2)

ods = Roo::OpenOffice.new("myspreadsheet.ods", password: "password")

ods.celltype
ods.comment(1,1, ods.sheets[-1])
ods.font(1,1).italic?
ods.formula('A', 2)

csv = Roo::CSV.new("mycsv.csv")

csv = Roo::CSV.new("mytsv.tsv", csv_options: {col_sep: "\t"})
csv = Roo::CSV.new("mycsv.csv", csv_options: {encoding: Encoding::ISO_8859_1})
```

```
```
