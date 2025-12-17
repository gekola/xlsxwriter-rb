# frozen_string_literal: true

require 'mkmf'
require 'English'

libxlsxwriter_dir = File.expand_path('libxlsxwriter', __dir__)
libxlsxwriter_src_dir = File.expand_path('src', libxlsxwriter_dir)
Dir.glob(File.expand_path('third_party/*', libxlsxwriter_dir)).each do |dir|
  make_pid = spawn({ 'CFLAGS' => $CFLAGS }, $make, chdir: dir)
  Process.wait(make_pid)
end
make_pid = spawn({ 'CFLAGS' => $CFLAGS }, $make, 'libxlsxwriter.a', chdir: libxlsxwriter_src_dir)
Process.wait(make_pid)
raise 'Make failed for xlsxwriter' unless $CHILD_STATUS.success?

append_cflags("-I'#{libxlsxwriter_dir}/include'")
append_ldflags('-lz')
append_ldflags("-L'#{libxlsxwriter_src_dir}'")
append_ldflags('-lxlsxwriter')

create_makefile 'xlsxwriter/xlsxwriter'
