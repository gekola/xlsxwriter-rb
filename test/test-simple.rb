require_relative './xlsx-func-testcase'

class TestSimple < XlsxWriterTestCase
  test 'simple01' do |wb|
    ws = wb.add_worksheet

    ws.write_string(0, 0, 'Hello', nil)
    ws.write_number(1, 0, 123,     nil)
  end

  test 'simple02' do |wb|
    ws1 = wb.add_worksheet
    ws2 = wb.add_worksheet('Data Sheet')
    ws3 = wb.add_worksheet

    wb.add_format(:bold, bold: true)

    ws1.write_string(0, 0, 'Foo', nil)
    ws1.write_number(1, 0, 123, nil)

    ws3.write_string(1, 1, 'Foo', nil)
    ws3.write_string(2, 1, 'Bar', :bold);
    ws3.write_number(3, 2, 234, nil);
  end

  test 'simple03' do |wb|
    ws1 = wb.add_worksheet
    ws2 = wb.add_worksheet("Data Sheet")
    ws3 = wb.add_worksheet

    wb.add_format(:bold, bold: true);

    ws1.write_string(0, 'A', 'Foo', nil)
    ws1.write_number(1, 'A', 123,   nil)

    ws3.write_string(1, 'B', 'Foo', nil)
    ws3.write_string(2, 'B', 'Bar', :bold)
    ws3.write_number(3, 'C4', 234,  nil)

    # Ensure the active worksheet is overwritten, below.
    x = ws2.activate

    y = ws2.select
    y = ws3.select
    y = ws3.activate
  end

  test 'simple04' do |wb|
    datetime1 = DateTime.new(0,    1,  1, 12, 0, 0)
    datetime2 = DateTime.new(2013, 1, 27,  0, 0, 0)

    ws = wb.add_worksheet

    wb.add_format(:f1, num_format_index: 20)
    wb.add_format(:f2, num_format_index: 14)

    ws.set_column(0, 0, width: 12)

    ws.write_datetime(0, 0, datetime1, :f1)
    ws.write_datetime(1, 0, datetime2, :f2)
  end
end
