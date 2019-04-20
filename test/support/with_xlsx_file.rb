require 'xlsxwriter'

module WithXlsxFile
  def with_xlsx_file(file_path = 'tmp/test.xlsx', **opts)
    after = opts.delete :after
    XlsxWriter::Workbook.open(file_path, opts) do |wb|
      yield wb
    end
    after.call if after
  ensure
    File.delete file_path if File.exist? file_path
  end
end
