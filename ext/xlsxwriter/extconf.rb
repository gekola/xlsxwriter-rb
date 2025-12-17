# frozen_string_literal: true

require 'mkmf'
require 'English'

libxlsxwriter_dir = File.expand_path('libxlsxwriter', __dir__)
libxlsxwriter_src_dir = File.join(libxlsxwriter_dir, 'src')
Dir.glob(File.join(libxlsxwriter_dir, 'third_party', '*')).each do |dir|
  make_pid = spawn({ 'CFLAGS' => $CFLAGS }, $make, chdir: dir)
  Process.wait(make_pid)
end
make_pid = spawn({ 'CFLAGS' => $CFLAGS }, $make, 'libxlsxwriter.a', chdir: libxlsxwriter_src_dir)
Process.wait(make_pid)
raise 'Make failed for xlsxwriter' unless $CHILD_STATUS.success?

find_header('xlsxwriter.h', File.join(libxlsxwriter_dir, 'include'))
have_library('z', 'crc32')
find_library('xlsxwriter', 'lxw_version', libxlsxwriter_src_dir)

create_makefile 'xlsxwriter/xlsxwriter'
