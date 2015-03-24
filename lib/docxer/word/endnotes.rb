module Docxer
  module Word
    class Endnotes

      attr_accessor :options
      def initialize(options={})
        @options = options
      end

      def render(zip)
        zip.put_next_entry('word/endnotes.xml')
        zip.write(Docxer.to_xml(document))
      end

      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.endnotes( 'xmlns:wpc' => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", 'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:wp14' => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", 'xmlns:wp' => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:wpg' => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", 'xmlns:wpi' => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", 'xmlns:wne' => "http://schemas.microsoft.com/office/word/2006/wordml", 'xmlns:wps' => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", 'mc:Ignorable' => "w14 wp14" ) do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
            xml['w'].endnote( 'w:type' => "separator", 'w:id' => "-1") do
              xml['w'].p do
                xml['w'].r do
                  xml['w'].separator
                end
              end
            end
            xml['w'].endnote( 'w:type' => "continuationSeparator", 'w:id' => "0" ) do
              xml['w'].p do
                xml['w'].r do
                  xml['w'].continuationSeparator
                end
              end
            end
          end
        end
      end

    end
  end
end