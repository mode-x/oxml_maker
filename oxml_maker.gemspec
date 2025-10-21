# frozen_string_literal: true

require_relative "lib/oxml_maker/version"

Gem::Specification.new do |spec|
  spec.name = "oxml_maker"
  spec.version = OxmlMaker::VERSION
  spec.authors = ["airbearr"]
  spec.email = ["emmanuel.abia@corsearch.com"]

  spec.summary = "Generate Microsoft Word DOCX files using OpenXML in Ruby"
  spec.description = "A Ruby gem for creating Microsoft Word DOCX files programmatically. Features include tables with dynamic data, paragraphs, custom page settings, Rails integration, and ZIP-based DOCX structure using rubyzip."
  spec.homepage = "https://github.com/mode-x/oxml_maker"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.1.0"

  spec.metadata["allowed_push_host"] = "https://rubygems.org"

  spec.metadata["homepage_uri"] = spec.homepage
  spec.metadata["source_code_uri"] = "https://github.com/mode-x/oxml_maker"
  spec.metadata["changelog_uri"] = "https://github.com/mode-x/oxml_maker/blob/main/CHANGELOG.md"

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  gemspec = File.basename(__FILE__)
  spec.files = IO.popen(%w[git ls-files -z], chdir: __dir__, err: IO::NULL) do |ls|
    ls.readlines("\x0", chomp: true).reject do |f|
      (f == gemspec) ||
        f.start_with?(*%w[bin/ test/ spec/ features/ .git appveyor Gemfile])
    end
  end
  spec.bindir = "exe"
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  # Uncomment to register a new dependency of your gem
  spec.add_dependency "rubyzip", "~> 3.2"

  # For more information and examples about making a new gem, check out our
  # guide at: https://bundler.io/guides/creating_gem.html
  spec.metadata["rubygems_mfa_required"] = "true"
end
