require_relative './xlsx-func-testcase'

class TestProperties < XlsxWriterTestCase
  test 'properties01' do |wb|
    wb.properties do |p|
      p.title    = 'This is an example spreadsheet'
      p.subject  = 'With document properties'
      p.author   = 'Someone'
      p.manager  = 'Dr. Heinz Doofenshmirtz'
      p.company  = 'of Wolves'
      p.category = 'Example spreadsheets'
      p.keywords = 'Sample, Example, Properties'
      p.comments = 'Created with Perl and Excel::Writer::XLSX'
      p.status   = 'Quo'
    end

    wb.add_worksheet do |ws|
      ws.set_column(0, 0, width: 70)
      ws.write_string(0, 0, 'Select \'Office Button -> Prepare -> Properties\' to see the file properties.', nil)
    end
  end

  test 'properties02' do |wb|
    wb.properties.hyperlink_base = 'C:\\'
    wb.add_worksheet
  end

  test 'properties04' do |wb|
    wb.properties['Checked by'] = 'Adam'
    wb.properties['Date completed'] = Time.new(2016, 12, 12, 23)
    wb.properties['Document number'] = 12345
    wb.properties['Reference'] = 1.2345
    wb.properties['Source'] = true
    wb.properties['Status'] = false
    wb.properties['Department'] = 'Finance'
    wb.properties['Group'] = 1.2345678901234

    wb.add_worksheet do |ws|
      ws.set_column 0, 0, width: 70
      ws.write_string 0, 'A', 'Select \'Office Button -> Prepare -> Properties\' to see the file properties.', nil
    end
  end

  test 'properties05' do |wb|
    wb.properties['Location'] = "Caf\xc3\xa9"
    wb.add_worksheet do |ws|
      ws.set_column 0, 0, width: 70
      ws.write_string 0, 'A', 'Select \'Office Button -> Prepare -> Properties\' to see the file properties.', nil
    end
  end
end
