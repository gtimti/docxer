module Docxer
  module Word
    class Fonts
      class Font

        attr_accessor :name, :options
        def initialize(name, options)
          @name = name
          @options = options
        end

        def build(xml)
          xml['w'].font('w:name' => @name) do
            xml['w'].panose1('w:val' => @options[:panose] )
            xml['w'].charset('w:val' => @options[:charset] )
            xml['w'].family('w:val' => @options[:family] )
            xml['w'].pitch('w:val' => @options[:pitch] )
            xml['w'].sig( @options[:sig] )
          end
        end

      end
    end
  end
end