#encoding: utf-8
module Docxer
  module Word
    class Numbering

      attr_accessor :options
      def initialize(options={})
        @options = options
      end

      def render(zip)
        zip.put_next_entry('word/numbering.xml')
        zip.write(Docxer.to_xml(document))
      end
      
      private
      def document
        Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
          xml.numbering( 'xmlns:wpc' => "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", 'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006", 'xmlns:o' => "urn:schemas-microsoft-com:office:office", 'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships", 'xmlns:m' => "http://schemas.openxmlformats.org/officeDocument/2006/math", 'xmlns:v' => "urn:schemas-microsoft-com:vml", 'xmlns:wp14' => "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", 'xmlns:wp' => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", 'xmlns:w10' => "urn:schemas-microsoft-com:office:word", 'xmlns:w' => "http://schemas.openxmlformats.org/wordprocessingml/2006/main", 'xmlns:w14' => "http://schemas.microsoft.com/office/word/2010/wordml", 'xmlns:wpg' => "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup", 'xmlns:wpi' => "http://schemas.microsoft.com/office/word/2010/wordprocessingInk", 'xmlns:wne' => "http://schemas.microsoft.com/office/word/2006/wordml", 'xmlns:wps' => "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", 'mc:Ignorable' => "w14 wp14") do
            xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "w" }
            xml['w'].abstractNum( 'w:abstractNumId' => 0 ) do
              xml['w'].nsid( 'w:val' => "7C8E1E67" )
              xml['w'].multiLevelType( 'w:val' => "hybridMultilevel" )
              xml['w'].tmpl( 'w:val' => "FFCCD426" )
              xml['w'].lvl( 'w:ilvl' => 0 ) do
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "-" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => 720, 'w:hanging' => 360 )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "1", 'w:tplc' => "04190003", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "o" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "1440", 'w:hanging' => 360 )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "2", 'w:tplc' => "04190005", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "2160", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "3", 'w:tplc' => "04190001", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "2880", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "4", 'w:tplc' => "04190003", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "o" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "3600", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "5", 'w:tplc' => "04190005", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "4320", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "6", 'w:tplc' => "04190001", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" ) 
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "5040", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "7", 'w:tplc' => "04190003", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "o" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "5760", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
              xml['w'].lvl( 'w:ilvl' => "8", 'w:tplc' => "04190005", 'w:tentative' => "1" ) do
                xml['w'].start( 'w:val' => "1" )
                xml['w'].numFmt( 'w:val' => "bullet" )
                xml['w'].lvlText( 'w:val' => "" )
                xml['w'].lvlJc( 'w:val' => "left" )
                xml['w'].pPr do
                  xml['w'].ind( 'w:left' => "6480", 'w:hanging' => "360" )
                end
                xml['w'].rPr do
                  xml['w'].rFonts( 'w:ascii' => "Cambria", 'w:eastAsia' => "Cambria", 'w:hAnsi' => "Cambria", 'w:cs' => "Times New Roman", 'w:hint' => "default" )
                end
              end
            end
            xml['w'].num( 'w:numId' => "1" ) do
              xml['w'].abstractNumId( 'w:val' => "0" )
            end
          end
        end
      end
    end
  end
end