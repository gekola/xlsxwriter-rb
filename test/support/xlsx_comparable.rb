# frozen_string_literal: true

require 'test/unit'
require 'zip'

module XlsxComparable
  def assert_xlsx_equal(got_path, exp_path, ignore_files=[], ignore_elements={})
    Zip::File.open(exp_path) do |exp_zip|
      Zip::File.open(got_path) do |got_zip|
        exp_files = exp_zip.glob('**/*').select { |e| !ignore_files.include?(e.name) }
        got_files = got_zip.glob('**/*').select { |e| !ignore_files.include?(e.name) }
        assert_equal(got_files.map(&:name).sort, exp_files.map(&:name).sort)

        exp_files.each do |exp_entry|
          next unless exp_entry.file?
          exp_xml_str = exp_zip.read(exp_entry.name)
          got_xml_str = got_zip.read(exp_entry.name)

          if %w(.png .jpeg .bmp .bin).include?(File.extname(exp_entry.name))
            exp_xml_str.force_encoding('BINARY')
            assert_equal(exp_xml_str, got_xml_str)
            next
          end

          case exp_entry.name
          when 'docProps/core.xml'
            exp_xml_str.gsub!(/ ?John/, '')
            exp_xml_str.gsub!(/\d{4}-\d\d-\d\dT\d\d\:\d\d:\d\dZ/, '')
            got_xml_str.gsub!(/\d{4}-\d\d-\d\dT\d\d\:\d\d:\d\dZ/, '')
          when 'xl/workbook.xml'
            exp_xml_str.gsub!(/<workbookView[^>]*>/, '<workbookView/>')
            exp_xml_str.gsub!(/<calcPr[^>]*>/, '<calcPr/>')
            got_xml_str.gsub!(/<workbookView[^>]*>/, '<workbookView/>')
            got_xml_str.gsub!(/<calcPr[^>]*>/, '<calcPr/>')
          when /xl\/worksheets\/sheet\d+.xml/
            exp_xml_str.gsub!(/horizontalDpi="200" /, '')
            exp_xml_str.gsub!(/verticalDpi="200" /, '')
            exp_xml_str.gsub!(/(<pageSetup[^>]*) r:id="rId1"/, '\1')
          when /xl\/charts\/chart\d+.xml/
            exp_xml_str.gsub!(/<c:pageMargins[^>]*>/, '<c:pageMargins/>')
            got_xml_str.gsub!(/<c:pageMargins[^>]*>/, '<c:pageMargins/>')
          end

          if exp_entry.name =~ /.vml\z/
            got_xml = _xml_to_list(got_xml_str)
            exp_xml = _vml_to_list(exp_xml_str)
          else
            got_xml = _xml_to_list(got_xml_str)
            exp_xml = _xml_to_list(exp_xml_str)
          end

          if ignore_elements.key?(exp_entry.name)
            patterns = ignore_elements[exp_entry.name]
            exp_xml.select! { |tag| patterns.none? { |pattern| tag.include? pattern } }
            got_xml.select! { |tag| patterns.none? { |pattern| tag.include? pattern } }
          end

          if exp_entry.name == '[Content_Types].xml' || exp_entry.name =~ /.rels\z/
            got_xml = _sort_rel_file_data(got_xml)
            exp_xml = _sort_rel_file_data(exp_xml)
          end

          assert_equal exp_xml, got_xml
        end
      end
    end
  end

  private

  def _xml_to_list(xml_str)
    xml_str.strip.split(/>\s*</).each do |el|
      el.gsub!("\r", '')
      el.insert 0, '<' unless el[0]  == '<'
      el <<        '>' unless el[-1] == '>'
    end
  end

  def _vml_to_list(vml_str)
    vml_str.gsub!("\r", '')

    vml = vml_str.split("\n")
    vml_str = String.new.tap do |vml_str|
      vml.each do |line|
        line.strip!
        next if line == ''
        line.gsub!(?', ?")
        line << ' ' if line =~ '"$'
        line << "\n" if line =~ '>$'
        line.gsub!('><', ">\n<")
        if line == "<x:Anchor>\n"
          line.strip!
        end
        wml_str << line
      end
    end
    vml_str.rstrip.split("\n")
  end

  def _sort_rel_file_data(xml_elements)
    first = xml_elements.shift
    last = xml_elements.pop

    xml_elements.sort!

    xml_elements.unshift(first)
    xml_elements.push(last)

    xml_elements
  end
end
