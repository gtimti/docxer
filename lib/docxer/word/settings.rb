module Docxer
  module Word
    class Settings

      attr_accessor :options
      def initialize(options={})
        @options = options
      end

      def render(zip)
        zip.put_next_entry('word/settings.xml')
        zip.write(Docxer.to_xml(document))
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.settings('xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:sl' => "http://schemas.openxmlformats.org/schemaLibrary/2006/main", 'mc:Ignorable' => "w14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
            xml['w'].zoom('w:percent' => 100)
            xml['w'].defaultTabStop('w:val' => 720)
            xml['w'].characterSpacingControl('w:val' => "doNotCompress")
            xml['w'].compat do
              xml['w'].compatSetting('w:name' => "compatibilityMode", 'w:uri' => "http://schemas.microsoft.com/office/word", 'w:val' => 14 )
              xml['w'].compatSetting('w:name' => "overrideTableStyleFontSizeAndJustification", 'w:uri' => "http://schemas.microsoft.com/office/word", 'w:val' => 1 )
              xml['w'].compatSetting('w:name' => "enableOpenTypeFeatures", 'w:uri' => "http://schemas.microsoft.com/office/word", 'w:val' => 1 )
              xml['w'].compatSetting('w:name' => "doNotFlipMirrorIndents", 'w:uri' => "http://schemas.microsoft.com/office/word", 'w:val' => 1 )
            end
            xml['w'].rsids do
              xml['w'].rsidRoot('w:val' => "00F6443A")
            end
            xml['m'].mathPr do
              xml['m'].mathFont('m:val' => "Cambria Math")
              xml['m'].brkBin('m:val' => "before")
              xml['m'].brkBinSub('m:val' => "--")
              xml['m'].smallFrac('m:val' => 0)
              xml['m'].dispDef
              xml['m'].lMargin('m:val' => 0)
              xml['m'].rMargin('m:val' => 0)
              xml['m'].defJc('m:val' => "centerGroup")
              xml['m'].wrapIndent('m:val' => 1440)
              xml['m'].intLim('m:val' => "subSup")
              xml['m'].naryLim('m:val' => "undOvr")
            end
            xml['w'].themeFontLang('w:val' => "en-US")
            xml['w'].clrSchemeMapping('w:bg1' => "light1", 'w:t1' => "dark1", 'w:bg2' => "light2", 'w:t2' => "dark2", 'w:accent1' => "accent1", 'w:accent2' => "accent2", 'w:accent3' => "accent3", 'w:accent4' => "accent4", 'w:accent5' => "accent5", 'w:accent6' => "accent6", 'w:hyperlink' => "hyperlink", 'w:followedHyperlink' => "followedHyperlink")
            xml['w'].shapeDefaults do
              xml['o'].shapedefaults( 'v:ext' => "edit", 'spidmax' => 1026 )
              xml['o'].shapelayout( 'v:ext' => "edit") do
                xml['o'].idmap('v:ext' => "edit", 'data' => 1)
              end
            end
            xml['w'].decimalSymbol('w:val' => ".")
            xml['w'].listSeparator('w:val' => ",")
          end
        end
      end

    end
  end
end
