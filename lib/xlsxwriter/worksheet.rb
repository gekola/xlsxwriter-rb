class XlsxWriter::Worksheet
  attr_reader :current_row

  THIN_CHARS = '^.acfijklrstxzFIJL()-'.freeze

  def add_row(row, style: nil, height: nil, types: nil)
    row_idx = @current_row ||= 0
    @current_row += 1

    row.each_with_index do |value, idx|
      cell_style = style.is_a?(Array) ? style[idx] : style
      cell_type = types.is_a?(Array) ? types[idx] : types

      case cell_type && cell_type.to_sym
      when :string
        write_string(row_idx, idx, value.to_s, cell_style)
        update_col_auto_width(idx, value.to_s, cell_style)
      when :number
        write_number(row_idx, idx, value.to_f, cell_style)
      when :formula
        write_formula(row_idx, idx, value, cell_style)
      when :datetime, :date, :time
        write_datetime(row_idx, idx, value.to_time, cell_style)
      when :url
        write_url(row_idx, idx, value, cell_style)
        update_col_auto_width(idx, value.to_s, cell_style)
      when :boolean
        write_boolean(row_idx, idx, value, cell_style)
      when :blank
        write_blank(row_idx, idx, cell_style)
      when nil
        case value
        when Numeric
          write_number(row_idx, idx, value, cell_style)
        when TrueClass, FalseClass
          write_boolean(row_idx, idx, value, cell_style)
        when '', nil
          write_blank(row_idx, idx, cell_style)
        when /\A=/
          write_formula(row_idx, idx, value, cell_style)
        else
          if value.respond_to? :to_time
            write_datetime(row_idx, idx, value.to_time, cell_style)
          else # assume string
            write_string(row_idx, idx, value.to_s, cell_style)
            update_col_auto_width(idx, value.to_s, cell_style)
          end
        end
      else
        raise ArgumentError, "Unknown cell type #{cell_type}."
      end

      set_row(row_idx, height: height) if height

      nil
    end
  end

  def update_col_auto_width(idx, val, format)
    font_scale = (@workbook.font_sizes[format] || 11) / 10.0
    width = (val.count(THIN_CHARS) + 3) * font_scale
    @col_auto_widths[idx] = [@col_auto_widths[idx], width].compact.max
  end

  def apply_auto_widths
    @col_auto_widths.each_with_index do |width, idx|
      set_column(idx, idx, width: width) if width
    end
  end
end
