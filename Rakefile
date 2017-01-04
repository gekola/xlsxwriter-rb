require 'rake/extensiontask'

task default: :test

spec = Gem::Specification.load('xlsxwriter.gemspec')
Rake::ExtensionTask.new('xlsxwriter', spec) do |ext|
  ext.lib_dir = 'lib/xlsxwriter'
end
task compile: :fetch_dep

Gem::PackageTask.new(spec) do |pkg|
end

DEP_DIR='ext/xlsxwriter/libxlsxwriter'

desc "Checkout xlsxwriter C library"
task :fetch_dep do
  rev = File.read('.dep_revision').chomp if File.exist?('.dep_revision')

  unless File.exist?('./ext/xlsxwriter/libxlsxwriter')
    `git clone https://github.com/jmcnamara/libxlsxwriter.git #{DEP_DIR}`
  else
    `cd #{DEP_DIR} && git fetch`
  end

  if rev
    `cd #{DEP_DIR} && git reset --hard #{rev}`
  else
    `(cd #{DEP_DIR} && git rev-parse HEAD) > .dep_revision`
  end

  Dir['./dep_patches/*.patch'].each do |patch|
    `(cd #{DEP_DIR} && patch -N -p1) <#{patch}`
  end
end

desc 'Run specs'
task test: :compile do
  ruby('test/run-test.rb')
end
