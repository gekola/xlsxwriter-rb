require_relative './xlsx-func-testcase'

class TestDefinedName < XlsxWriterTestCase
  test 'defined_name01' do |wb, t|
    t.ignore_elements = { 'xl/worksheets/sheet1.xml' => [ '<pageMargins' ] }
    wb.add_worksheet do |ws|
      ws.paper = 9
      ws.vertical_dpi = 200
      ws
        .print_area(0, 'A', 5, 'E')
        .autofilter(0, 'F', 0, 'G')
        .write_string(0, 'G', 'Filter')
        .write_string(0, 'F', 'Auto')
        .fit_to_pages(2, 2)
    end

    wb.add_worksheet
    wb.add_worksheet 'Sheet 3'

    wb.define_name "'Sheet 3'!Bar", "='Sheet 3'!$A$1"
    wb.define_name "Abc",           "=Sheet1!$A$1"
    wb.define_name "Baz",           "=0.98"
    wb.define_name "Sheet1!Bar",    "=Sheet1!$A$1"
    wb.define_name "Sheet2!Bar",    "=Sheet2!$A$1"
    wb.define_name "Sheet2!aaa",    "=Sheet2!$A$1"
    wb.define_name "_Egg",          "=Sheet1!$A$1"
    wb.define_name "_Fog",          "=Sheet1!$A$1"
  end
end
