require_relative './xlsx-func-testcase'

class TestDefaultRow < XlsxWriterTestCase
  test 'default_row01' do |wb|
    ws = wb.add_worksheet

    ws.set_default_row(24, false)

    ws.write_string(0, 'A', 'Foo', nil)
    ws.write_string(9, 'A', 'Bar', nil)
  end

  test 'default_row05' do |wb|
    ws = wb.add_worksheet

    ws.set_default_row(24, true)

    ws.write_string(0,  'A', 'Foo', nil)
    ws.write_string(9,  'A', 'Bar', nil)
    ws.write_string(19, 'A', 'Baz', nil)

    (1..8).each { |i| ws.set_row(i, height: 24) }
    (10..19).each { |i| ws.set_row(i, height: 24) }
  end
end
