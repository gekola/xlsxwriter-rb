# frozen_string_literal: true

require 'rake/extensiontask'

task default: :test

spec = Gem::Specification.load('xlsxwriter.gemspec')
Rake::ExtensionTask.new('xlsxwriter', spec) do |ext|
  ext.lib_dir = 'lib/xlsxwriter'
end
Rake::Task['compile'].prerequisites.unshift :patch_dep

Gem::PackageTask.new(spec) do |pkg|
end

DEP_DIR = 'ext/xlsxwriter/libxlsxwriter'

desc 'Checkout xlsxwriter C library'
task :patch_dep do
  patches = Dir["#{pwd}/dep_patches/*.patch"]
  chdir(DEP_DIR) do
    if File.exist?('.git')
      sh 'git reset --hard'

      patches.each do |patch|
        sh "patch -N -p1 <#{patch}"
      end
    end
  end
end

desc 'Run specs'
task test: :compile do
  ruby 'test/run_test.rb'
end
