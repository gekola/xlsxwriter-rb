# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestArrayFormula < XlsxWriterTestCase
  test 'array_formula01' do |wb, t|
    t.ignore_files = %w(
      xl/calcChain.xml
      [Content_Types].xml
      xl/_rels/workbook.xml.rels
    )
    ws = wb.add_worksheet nil

    ws.write_number(0, 1, 0, nil)
    ws.write_number(1, 1, 0, nil)
    ws.write_number(2, 1, 0, nil)
    ws.write_number(0, 2, 0, nil)
    ws.write_number(1, 2, 0, nil)
    ws.write_number(2, 2, 0, nil)

    ws.write_array_formula(0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}', nil)
  end

  test 'array_formula02' do |wb, t|
    t.ignore_files = %w(
      xl/calcChain.xml
      [Content_Types].xml
      xl/_rels/workbook.xml.rels
    )
    ws = wb.add_worksheet
    wb.add_format(:bold, bold: true)

    ws.write_number(0, 1, 0, nil)
    ws.write_number(1, 1, 0, nil)
    ws.write_number(2, 1, 0, nil)
    ws.write_number(0, 2, 0, nil)
    ws.write_number(1, 2, 0, nil)
    ws.write_number(2, 2, 0, nil)

    ws.write_array_formula(0, 0, 2, 0, '{=SUM(B1:C1*B2:C2)}', :bold)
  end
end
