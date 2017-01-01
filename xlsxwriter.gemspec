$LOAD_PATH << File.expand_path('../lib', __FILE__)

require 'xlsxwriter/version'

Gem::Specification.new do |s|
  s.name = 'xlsxwriter'.freeze
  s.version = XlsxWriter::VERSION
  s.summary = 'Ruby interface to libxlsxwriter'.freeze
  s.authors = ['Nick H'.freeze].freeze
  s.homepage = 'https://github.com/gekola/xlsxwriter-rb'
  s.license = 'BSD-2-Clause-FreeBSD'
  s.files = ['Rakefile',
             *Dir['{lib,test}/**/*.rb'],
             *Dir['ext/xlsxwriter/*.{c,h,rb}'],
             'ext/xlsxwriter/libxlsxwriter/LICENSE.txt',
             'ext/xlsxwriter/libxlsxwriter/Makefile',
             'ext/xlsxwriter/libxlsxwriter/lib/.gitignore',
             *Dir['ext/xlsxwriter/libxlsxwriter/{src,third_party}/**/Makefile'],
             *Dir['ext/xlsxwriter/libxlsxwriter/{src,third_party,include}/**/*.{c,h}']].map(&:freeze)
  s.extensions = Dir['ext/**/extconf.rb'].map(&:freeze)
  s.test_files = Dir['test/**/*'].map(&:freeze)

  s.add_development_dependency 'rake-compiler', '~>1.0'
  s.add_development_dependency 'test-unit', '~>3.2'
  s.add_development_dependency 'rubyzip', '~>1.2'
end
