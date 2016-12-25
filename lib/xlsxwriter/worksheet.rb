require 'date'
require 'uri'

class XlsxWriter::Worksheet
  attr_reader :current_row
  def add_row(row, style: nil, height: nil, types: nil)
    row_idx = @current_row ||= 0
    @current_row += 1

    row.each_with_index do |value, idx|
      cell_style = style.is_a?(Array) ? style[idx] : nil
      cell_type = types.is_a?(Array) ? types[idx] : types

      case cell_type && cell_type.to_sym
      when :string
        write_string(row_idx, idx, value.to_s, cell_style)
      when :number
        write_number(row_idx, idx, value, cell_style)
      when :formula
        write_formula(row_idx, idx, value, cell_style)
      when :datetime
        write_datetime(row_idx, idx, value.to_datetime, cell_style)
      when :url
        write_url(row_idx, idx, value, cell_style)
      when :boolean
        write_boolean(row_idx, idx, value, cell_style)
      when :blank
        write_blank(row_idx, idx, cell_style)
      when nil
        case value
        when Numeric
          write_number(row_idx, idx, value, cell_style)
        when DateTime, Date, Time
          write_datetime(row_idx, idx, value.to_datetime, cell_style)
        when TrueClass, FalseClass
          write_boolean(row_idx, idx, value, cell_style)
        when '', nil
          write_blank(row_idx, idx, cell_style)
        when /\A=/
          write_formula(row_idx, idx, value, cell_style)
        when URI.regexp
          write_url(row_idx, idx, value, cell_style)
        when String
          write_string(row_idx, idx, value, cell_style)
        else # assume string
          write_string(row_idx, idx, value.to_s, cell_style)
        end
      else
        raise ArgumentError, "Unknown cell type #{cell_type}."
      end

      nil
    end
  end
end
