# frozen_string_literal: true

require_relative './xlsx-func-testcase'

class TestFitToPages < XlsxWriterTestCase
  test 'fit_to_pages01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }

    ws = wb.add_worksheet
    ws.fit_to_pages 1, 1
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 'A1', 'Foo'
  end

  test 'fit_to_pages02' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }

    ws = wb.add_worksheet
    ws.fit_to_pages 2, 1
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 'A1', 'Foo'
  end

  test 'fit_to_pages03' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }

    ws = wb.add_worksheet
    ws.fit_to_pages 1, 2
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 'A1', 'Foo'
  end

  test 'fit_to_pages04' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }

    ws = wb.add_worksheet
    ws.fit_to_pages(3, 2)
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 0, 'A', 'Foo'
  end

  test 'fit_to_pages05' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }

    ws = wb.add_worksheet
    ws.fit_to_pages(1, 0)
    ws.paper = 9
    ws.vertical_dpi = 200
    ws.write_string 0, 'A', 'Foo'
  end
end
