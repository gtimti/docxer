require 'docxer/properties/app'
require 'docxer/properties/core'

module Docxer
  class Properties
    attr_accessor :options, :app, :core

    def initialize(options)
      @options = options
    end

    def app
      @app ||= App.new(@options)
    end

    def core
      @core ||= Core.new(@options)
    end

    def render(zip)
      zip.put_next_entry('docProps/app.xml')
      zip.write(Docxer.to_xml(app.document))

      zip.put_next_entry('docProps/core.xml')
      zip.write(Docxer.to_xml(core.document))
    end

  end
end