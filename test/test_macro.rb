# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestMacro < XlsxWriterTestCase
  test 'macro01.xlsm' do |wb|
    wb
      .add_vba_project(image_path('vbaProject01.bin'))
      .add_worksheet
      .write_number(0, 0, 123)
  end

  test 'macro02.xlsm' do |wb|
    wb.add_vba_project(image_path('vbaProject03.bin'))
    wb.vba_name = 'MyWorkbook'
    ws = wb.add_worksheet
    ws.vba_name = 'MySheet1'
    ws.write_number(0, 0, 123)
  end
end
