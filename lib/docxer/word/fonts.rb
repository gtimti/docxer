require 'docxer/word/fonts/font'

module Docxer
  module Word
    class Fonts

      attr_accessor :fonts

      def initialize
        @fonts = []
        add('Calibri', :panose => '020F0502020204030204', :charset => '00', :family => 'swiss', :pitch => 'variable', :sig => { 'w:usb0' => "E00002FF", 'w:usb1' => "4000ACFF", 'w:usb2' => "00000001", 'w:usb3' => "00000000", 'w:csb0' => "0000019F", 'w:csb1' => "00000000"})
        add('Times New Roman', :panose => '02020603050405020304', :charset => '00', :family => 'roman', :pitch => 'variable', :sig => { 'w:usb0' => "E0002AFF", 'w:usb1' => "C0007841", 'w:usb2' => "00000009", 'w:usb3' => "00000000", 'w:csb0' => "000001FF", 'w:csb1' => "00000000"})
        add('Cambria', :panose => '02040503050406030204', :charset => '00', :family => 'roman', :pitch => 'variable', :sig => { 'w:usb0' => "E00002FF", 'w:usb1' => "400004FF", 'w:usb2' => "00000000", 'w:usb3' => "00000000", 'w:csb0' => "0000019F", 'w:csb1' => "00000000"})
      end

      def add(name, options)
        font = Font.new(name, options)
        @fonts << font if font
        font
      end

      def render(zip)
        zip.put_next_entry('word/fontTable.xml')
        zip.write(Docxer.to_xml(document))
      end

      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.fonts('xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'mc:Ignorable' => "w14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix=="w" }
            
            @fonts.each do |font|
              font.build(xml)
            end
          end
        end
      end

    end
  end
end