#encoding: utf-8

module Docxer
  module Word
    class Themes
      class Theme

        attr_accessor :sequence
        def initialize
        end

        def set_sequence(sequence)
          @sequence = sequence
          self
        end

        def render
          Docxer.to_xml(document)
        end

        private
        def document
          Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
            xml.theme( 'xmlns:a' => "http://schemas.openxmlformats.org/drawingml/2006/main", 'name' => "Office Theme" ) do
              xml.parent.namespace = xml.parent.namespace_definitions.find{|ns| ns.prefix == "a" }
              xml['a'].themeElements do
                xml['a'].clrScheme( 'name' => "Office") do
                  xml['a'].dk1 do
                    xml['a'].sysClr( 'val' => "windowText", 'lastClr' => "000000" )
                  end
                  xml['a'].lt1 do
                    xml['a'].sysClr( 'val' => "window", 'lastClr' => "FFFFFF" )
                  end
                  xml['a'].dk2 do
                    xml['a'].srgbClr( 'val' => "1F497D" )
                  end
                  xml['a'].lt2 do
                    xml['a'].srgbClr( 'val' => "EEECE1" )
                  end
                  xml['a'].accent1 do
                    xml['a'].srgbClr( 'val' => "4F81BD" )
                  end
                  xml['a'].accent2 do
                    xml['a'].srgbClr( 'val' => "C0504D" )
                  end
                  xml['a'].accent3 do
                    xml['a'].srgbClr( 'val' => "9BBB59" )
                  end
                  xml['a'].accent4 do
                    xml['a'].srgbClr( 'val' => "8064A2" )
                  end
                  xml['a'].accent5 do
                    xml['a'].srgbClr( 'val' => "4BACC6" )
                  end
                  xml['a'].accent6 do
                    xml['a'].srgbClr( 'val' => "F79646" )
                  end
                  xml['a'].hlink do
                    xml['a'].srgbClr( 'val' => "0000FF" )
                  end
                  xml['a'].folHlink do
                    xml['a'].srgbClr( 'val' => "800080" )
                  end
                end
                xml['a'].fontScheme( 'name' => "Office" ) do
                  xml['a'].majorFont do
                    xml['a'].latin( 'typeface' => "Cambria" )
                    xml['a'].ea( 'typeface' => "" )
                    xml['a'].cs( 'typeface' => "" )
                    xml['a'].font( 'script' => "Jpan", 'typeface' => "ＭＳ ゴシック" )
                    xml['a'].font( 'script' => "Hang", 'typeface' => "맑은 고딕" )
                    xml['a'].font( 'script' => "Hans", 'typeface' => "宋体" )
                    xml['a'].font( 'script' => "Hant", 'typeface' => "新細明體" )
                    xml['a'].font( 'script' => "Arab", 'typeface' => "Times New Roman" )
                    xml['a'].font( 'script' => "Hebr", 'typeface' => "Times New Roman")
                    xml['a'].font( 'script' => "Thai", 'typeface' => "Angsana New")
                    xml['a'].font( 'script' => "Ethi", 'typeface' => "Nyala")
                    xml['a'].font( 'script' => "Beng", 'typeface' => "Vrinda")
                    xml['a'].font( 'script' => "Gujr", 'typeface' => "Shruti")
                    xml['a'].font( 'script' => "Khmr", 'typeface' => "MoolBoran")
                    xml['a'].font( 'script' => "Knda", 'typeface' => "Tunga")
                    xml['a'].font( 'script' => "Guru", 'typeface' => "Raavi")
                    xml['a'].font( 'script' => "Cans", 'typeface' => "Euphemia")
                    xml['a'].font( 'script' => "Cher", 'typeface' => "Plantagenet Cherokee")
                    xml['a'].font( 'script' => "Yiii", 'typeface' => "Microsoft Yi Baiti")
                    xml['a'].font( 'script' => "Tibt", 'typeface' => "Microsoft Himalaya")
                    xml['a'].font( 'script' => "Thaa", 'typeface' => "MV Boli")
                    xml['a'].font( 'script' => "Deva", 'typeface' => "Mangal")
                    xml['a'].font( 'script' => "Telu", 'typeface' => "Gautami")
                    xml['a'].font( 'script' => "Taml", 'typeface' => "Latha")
                    xml['a'].font( 'script' => "Syrc", 'typeface' => "Estrangelo Edessa")
                    xml['a'].font( 'script' => "Orya", 'typeface' => "Kalinga")
                    xml['a'].font( 'script' => "Mlym", 'typeface' => "Kartika")
                    xml['a'].font( 'script' => "Laoo", 'typeface' => "DokChampa")
                    xml['a'].font( 'script' => "Sinh", 'typeface' => "Iskoola Pota")
                    xml['a'].font( 'script' => "Mong", 'typeface' => "Mongolian Baiti")
                    xml['a'].font( 'script' => "Viet", 'typeface' => "Times New Roman")
                    xml['a'].font( 'script' => "Uigh", 'typeface' => "Microsoft Uighur")
                    xml['a'].font( 'script' => "Geor", 'typeface' => "Sylfaen")
                  end
                  xml['a'].minorFont do
                    xml['a'].latin( 'typeface' => "Calibri")
                    xml['a'].ea( 'typeface' => "")
                    xml['a'].cs( 'typeface' => "")
                    xml['a'].font( 'script' => "Jpan", 'typeface' => "ＭＳ 明朝")
                    xml['a'].font( 'script' => "Hang", 'typeface' => "맑은 고딕")
                    xml['a'].font( 'script' => "Hans", 'typeface' => "宋体")
                    xml['a'].font( 'script' => "Hant", 'typeface' => "新細明體")
                    xml['a'].font( 'script' => "Arab", 'typeface' => "Arial")
                    xml['a'].font( 'script' => "Hebr", 'typeface' => "Arial")
                    xml['a'].font( 'script' => "Thai", 'typeface' => "Cordia New")
                    xml['a'].font( 'script' => "Ethi", 'typeface' => "Nyala")
                    xml['a'].font( 'script' => "Beng", 'typeface' => "Vrinda")
                    xml['a'].font( 'script' => "Gujr", 'typeface' => "Shruti")
                    xml['a'].font( 'script' => "Khmr", 'typeface' => "DaunPenh")
                    xml['a'].font( 'script' => "Knda", 'typeface' => "Tunga")
                    xml['a'].font( 'script' => "Guru", 'typeface' => "Raavi")
                    xml['a'].font( 'script' => "Cans", 'typeface' => "Euphemia")
                    xml['a'].font( 'script' => "Cher", 'typeface' => "Plantagenet Cherokee")
                    xml['a'].font( 'script' => "Yiii", 'typeface' => "Microsoft Yi Baiti")
                    xml['a'].font( 'script' => "Tibt", 'typeface' => "Microsoft Himalaya")
                    xml['a'].font( 'script' => "Thaa", 'typeface' => "MV Boli")
                    xml['a'].font( 'script' => "Deva", 'typeface' => "Mangal")
                    xml['a'].font( 'script' => "Telu", 'typeface' => "Gautami")
                    xml['a'].font( 'script' => "Taml", 'typeface' => "Latha")
                    xml['a'].font( 'script' => "Syrc", 'typeface' => "Estrangelo Edessa")
                    xml['a'].font( 'script' => "Orya", 'typeface' => "Kalinga")
                    xml['a'].font( 'script' => "Mlym", 'typeface' => "Kartika")
                    xml['a'].font( 'script' => "Laoo", 'typeface' => "DokChampa")
                    xml['a'].font( 'script' => "Sinh", 'typeface' => "Iskoola Pota")
                    xml['a'].font( 'script' => "Mong", 'typeface' => "Mongolian Baiti")
                    xml['a'].font( 'script' => "Viet", 'typeface' => "Arial")
                    xml['a'].font( 'script' => "Uigh", 'typeface' => "Microsoft Uighur")
                    xml['a'].font( 'script' => "Geor", 'typeface' => "Sylfaen")
                  end
                end
                xml['a'].fmtScheme( 'name' => "Office") do
                  xml['a'].fillStyleLst do
                    xml['a'].solidFill do
                      xml['a'].schemeClr( 'val' => "phClr" )
                    end
                    xml['a'].gradFill( 'rotWithShape' => "1") do
                      xml['a'].gsLst do
                        xml['a'].gs( 'pos' => "0" ) do
                          xml['a'].schemeClr( 'val' => "phClr" ) do
                            xml['a'].tint( 'val' => "50000" )
                            xml['a'].satMod( 'val' => "300000" )
                          end
                        end
                        xml['a'].gs( 'pos' => "35000") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].tint( 'val' => "37000")
                            xml['a'].satMod( 'val' => "300000") 
                          end
                        end
                        xml['a'].gs( 'pos' => "100000" ) do
                          xml['a'].schemeClr( 'val' => "phClr" ) do
                            xml['a'].tint( 'val' => "15000")
                            xml['a'].satMod( 'val' => "350000")
                          end
                        end
                      end
                      xml['a'].lin( 'ang' => "16200000", 'scaled' => "1")
                    end
                    xml['a'].gradFill( 'rotWithShape' => "1") do
                      xml['a'].gsLst do
                        xml['a'].gs( 'pos' => "0") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].shade( 'val' => "51000")
                            xml['a'].satMod( 'val' => "130000")
                          end
                        end
                        xml['a'].gs( 'pos' => "80000" ) do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].shade( 'val' => "93000")
                            xml['a'].satMod( 'val' => "130000")
                          end
                        end
                        xml['a'].gs( 'pos' => "100000") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].shade( 'val' => "94000")
                            xml['a'].satMod( 'val' => "135000")
                          end
                        end
                      end
                      xml['a'].lin( 'ang' => "16200000", 'scaled' => "0")
                    end
                  end
                  xml['a'].lnStyleLst do
                    xml['a'].ln( 'w' => "9525", 'cap' => "flat", 'cmpd' => "sng", 'algn' => "ctr") do
                      xml['a'].solidFill do
                        xml['a'].schemeClr( 'val' => "phClr") do
                          xml['a'].shade( 'val' => "95000")
                          xml['a'].satMod( 'val' => "105000")
                        end
                      end
                      xml['a'].prstDash( 'val' => "solid")
                    end
                    xml['a'].ln( 'w' => "25400", 'cap' => "flat", 'cmpd' => "sng", 'algn' => "ctr") do
                      xml['a'].solidFill do
                        xml['a'].schemeClr( 'val' => "phClr")
                      end
                      xml['a'].prstDash( 'val' => "solid")
                    end
                    xml['a'].ln( 'w' => "38100", 'cap' => "flat", 'cmpd' => "sng", 'algn' => "ctr") do
                      xml['a'].solidFill do
                        xml['a'].schemeClr( 'val' => "phClr")
                      end
                      xml['a'].prstDash( 'val' => "solid")
                    end
                  end
                  xml['a'].effectStyleLst do
                    xml['a'].effectStyle do
                      xml['a'].effectLst do
                        xml['a'].outerShdw( 'blurRad' => "40000", 'dist' => "20000", 'dir' => "5400000", 'rotWithShape' => "0") do
                          xml['a'].srgbClr( 'val' => "000000") do
                            xml['a'].alpha( 'val' => "38000")
                          end
                        end
                      end
                    end
                    xml['a'].effectStyle do
                      xml['a'].effectLst do
                        xml['a'].outerShdw( 'blurRad' => "40000", 'dist' => "23000", 'dir' => "5400000", 'rotWithShape' => "0") do
                          xml['a'].srgbClr( 'val' => "000000") do
                            xml['a'].alpha( 'val' => "35000")
                          end
                        end
                      end
                    end
                    xml['a'].effectStyle do
                      xml['a'].effectLst do
                        xml['a'].outerShdw( 'blurRad' => "40000", 'dist' => "23000", 'dir' => "5400000", 'rotWithShape' => "0") do
                          xml['a'].srgbClr( 'val' => "000000") do
                            xml['a'].alpha( 'val' => "35000")
                          end
                        end
                      end
                      xml['a'].scene3d do
                        xml['a'].camera( 'prst' => "orthographicFront") do
                          xml['a'].rot( 'lat' => "0", 'lon' => "0", 'rev' => "0")
                        end
                        xml['a'].lightRig( 'rig' => "threePt", 'dir' => "t") do
                          xml['a'].rot( 'lat' => "0", 'lon' => "0", 'rev' => "1200000")
                        end
                      end
                      xml['a'].sp3d do
                        xml['a'].bevelT( 'w' => "63500", 'h' => "25400")
                      end
                    end
                  end
                  xml['a'].bgFillStyleLst do
                    xml['a'].solidFill do
                      xml['a'].schemeClr( 'val' => "phClr")
                    end
                    xml['a'].gradFill( 'rotWithShape' => "1") do
                      xml['a'].gsLst do
                        xml['a'].gs( 'pos' => "0") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].tint( 'val' => "40000")
                            xml['a'].satMod( 'val' => "350000")
                          end
                        end
                        xml['a'].gs( 'pos' => "40000") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].tint( 'val' => "45000")
                            xml['a'].shade( 'val' => "99000")
                            xml['a'].satMod( 'val' => "350000")
                          end
                        end
                        xml['a'].gs( 'pos' => "100000") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].shade( 'val' => "20000")
                            xml['a'].satMod( 'val' => "255000")
                          end
                        end
                      end
                      xml['a'].path( 'path' => "circle") do
                        xml['a'].fillToRect( 'l' => "50000", 't' => "-80000", 'r' => "50000", 'b' => "180000")
                      end
                    end
                    xml['a'].gradFill( 'rotWithShape' => "1") do
                      xml['a'].gsLst do
                        xml['a'].gs( 'pos' => "0") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].tint( 'val' => "80000")
                            xml['a'].satMod( 'val' => "300000")
                          end
                        end
                        xml['a'].gs( 'pos' => "100000") do
                          xml['a'].schemeClr( 'val' => "phClr") do
                            xml['a'].shade( 'val' => "30000")
                            xml['a'].satMod( 'val' => "200000")
                          end
                        end
                      end
                      xml['a'].path( 'path' => "circle") do
                        xml['a'].fillToRect( 'l' => "50000", 't' => "50000", 'r' => "50000", 'b' => "50000")
                      end
                    end
                  end
                end
              end
              xml['a'].objectDefaults
              xml['a'].extraClrSchemeLst
            end
          end
        end

      end
    end
  end
end