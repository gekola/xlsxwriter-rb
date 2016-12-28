require_relative './xlsx-func-testcase'

class TestData < XlsxWriterTestCase
  test 'data01' do |wb|
    ws = wb.add_worksheet
    ws.write_string(0, 0, 'Hello', nil)
  end

  test 'data02' do |wb|
    ws = wb.add_worksheet

    # in range
    ws.write_number(0, 0, 123, nil)
    ws.write_number(1_048_575, 0, 456, nil)

    # out of range
    ws.write_number(-1, 0, 123, nil)
    ws.write_number(1_048_576, 0, 456, nil)
  end

  test 'data03' do |wb|
    ws = wb.add_worksheet

    ws.write_number(0, 16_383, 123, nil)
    ws.write_number(1_048_575, 16_383, 456, nil)
  end

  test 'data04' do |wb|
    ws = wb.add_worksheet

    ws.write_string(0, 0, 'Foo', nil)
    ws.write_string(0, 1, 'Bar', nil)
    ws.write_string(1, 0, 'Bing', nil)
    ws.write_string(2, 0, 'Buzz', nil)
    ws.write_string(1_048_575, 0, 'End', nil)
  end

  test 'data05' do |wb|
    ws = wb.add_worksheet

    wb.add_format :format, bold: 1
    ws.write_string 0, 0, 'Foo', :format
  end

  test 'data06' do |wb|
    ws = wb.add_worksheet

    wb.add_format :f1, bold: true
    wb.add_format :f2, italic: true
    wb.add_format :f3, bold: true, italic: true

    ws.write_string 0, 'A', 'Foo', :f1
    ws.write_string 1, 'A', 'Bar', :f2
    ws.write_string 2, 'A', 'Baz', :f3
  end

  test 'data07' do |wb, t|
    t.ignore_files = ['xl/calcChain.xml',
                      '[Content_Types].xml',
                      'xl/_rels/workbook.xml.rels']
    ws = wb.add_worksheet

    ws.write_formula_num(0, 0, '=1+2', nil, 3)
  end
end
