# frozen_string_literal: true

require_relative './xlsx_func_testcase'

class TestSimple < XlsxWriterTestCase
  test 'simple01' do |wb|
    ws = wb.add_worksheet

    ws.write_string(0, 0, 'Hello')
    ws.write_number(1, 0, 123)
  end

  test 'simple02' do |wb|
    ws1 = wb.add_worksheet
    wb.add_worksheet('Data Sheet')
    ws3 = wb.add_worksheet

    wb.add_format(:bold, bold: true)

    ws1.write_string(0, 0, 'Foo')
    ws1.write_number(1, 0, 123)

    ws3.write_string('1', '1', 'Foo')
    ws3.write_string('2', 'B', 'Bar', :bold)
    ws3.write_number('3', 2, 234)
  end

  test 'simple03' do |wb|
    ws1 = wb.add_worksheet
    ws2 = wb.add_worksheet('Data Sheet')
    ws3 = wb.add_worksheet

    wb.add_format(:bold, bold: true)

    ws1.write_string(0, 'A', 'Foo')
    ws1.write_number(1, 'A', 123)

    ws3.write_string(1, 'B', 'Foo')
    ws3.write_string(2, 'B', 'Bar', :bold)
    ws3.write_number('C4', 234)

    # Ensure the active worksheet is overwritten, below.
    ws2.activate

    ws2.select
    ws3.select
    ws3.activate
  end

  test 'simple04' do |wb|
    datetime1 = Time.new(0,    1, 1, 12)
    datetime2 = Time.new(2013, 1, 27, 0)

    ws = wb.add_worksheet

    wb.add_format(:f1, num_format_index: 20)
    wb.add_format(:f2, num_format_index: 14)

    ws.set_column(0, 0, width: 12)

    ws.write_datetime('A1', datetime1, :f1)
    ws.write_datetime(1, 0, datetime2, :f2)
  end
end
