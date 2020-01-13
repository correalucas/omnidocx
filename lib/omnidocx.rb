require 'omnidocx/version'
require 'nokogiri'
require 'zip'
require 'tempfile'
require 'mime/types'
require 'open-uri'

class Omnidocx::Docx
  RELATIONSHIP_FILE_PATH = 'word/_rels/document.xml.rels'.freeze
  CONTENT_TYPES_FILE = '[Content_Types].xml'.freeze
  STYLES_FILE_PATH = 'word/styles.xml'.freeze

  NAMESPACES = {
    "w": 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    "wp": 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    "a": 'http://schemas.openxmlformats.org/drawingml/2006/main',
    "pic": 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    "r": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  }.freeze

  def self.document_file(file)
    rel_xml = Nokogiri::XML(file.read('_rels/.rels'))
    rel_xml.at_css('[Id="rId1"]').attr('Target').gsub(%r{^\/}, '')
  end

  def self.merge_documents(documents_to_merge = [], merge_place = nil, final_path)
    temp_file = Tempfile.new('docxedit-')
    documents_to_merge_count = documents_to_merge.count

    # minimum two documents required to merge
    return 'Pass at least two documents to be merged' if documents_to_merge_count < 2

    # first document to which the others will be appended (header/footer will be picked from this document)
    @main_document_zip = Zip::File.new(documents_to_merge.first)
    document_file_path = document_file(@main_document_zip)
    @main_document_xml = Nokogiri::XML(@main_document_zip.read(document_file_path))
    @main_body = @main_document_xml.xpath('//w:body')

    @place = @main_body.children.last
    @place = @main_body.children.xpath("//w:t[contains(text(),'#{merge_place}')]").first if merge_place.present?

    return 'Place to merge document not found' if @place.blank?

    @rel_doc = ''
    @cont_type_doc = ''
    @style_doc = ''
    doc_cnt = 0
    # cnt variable to construct relationship ids, taken a high value 100 to avoid duplication
    cnt = 100
    tbl_cnt = 10
    # hash to store information about the media files and their corresponding new names
    media_hash = {}
    # rid_hash to store relationship information
    rid_hash = {}
    # table hash to store information if any tables present
    table_hash = {}
    # head_foot_media hash to store if any media files present in header/footer
    head_foot_media = {}
    # a counter for docPr element in the main document body
    doc_pr_id = 100

    # array to store content type information about media extensions
    default_extensions = []
    # array to store override content type information
    override_partnames = []

    # array to store information about additional content types other than the ones present in the first(main) document
    additional_cont_type_entries = []

    #  prepare initial set of data from first document
    @main_document_zip.entries.each do |zip_entrie|
      in_stream = zip_entrie.get_input_stream.read

      # Relationship XML
      @rel_doc = Nokogiri::XML(in_stream) if zip_entrie.name == RELATIONSHIP_FILE_PATH

      # Styles XML to be updated later on with the additional tables info
      @style_doc = Nokogiri::XML(in_stream) if zip_entrie.name == STYLES_FILE_PATH

      # Content types XML to be updated later on with the additional media type info
      if zip_entrie.name == CONTENT_TYPES_FILE
        @cont_type_doc = Nokogiri::XML in_stream
        default_nodes = @cont_type_doc.css "Default"
        override_nodes = @cont_type_doc.css "Override"
        default_nodes.each { |node| default_extensions << node["Extension"] }
        override_nodes.each { |node| override_partnames << node["PartName"] }
      end
    end

    # opening a new zip for the final document
    Zip::OutputStream.open(temp_file.path) do |zos|
      documents_to_merge.each do |doc_path|
        media_hash["doc#{doc_cnt}"] = {}
        rid_hash["doc#{doc_cnt}"] = {}
        head_foot_media["doc#{doc_cnt}"] = []
        table_hash["doc#{doc_cnt}"] = {}
        zip_file = Zip::File.new(doc_path)

        zip_file.entries.each do |e|
          if ['word/_rels/header', 'word/_rels/footer'].include?(e.name)
            hf_xml = Nokogiri::XML(e.get_input_stream.read)
            hf_xml.css("Relationship").each do |rel_node|
              # media file names in header & footer need not be changed as they will be picked from the first document only and not the subsequent documents, so no chance of duplication
              head_foot_media["doc#{doc_cnt}"] << rel_node["Target"].gsub("media/", "")
            end
          end
          if e.name == CONTENT_TYPES_FILE
            cont_type_xml = Nokogiri::XML(e.get_input_stream.read)
            default_nodes = cont_type_xml.css "Default"
            override_nodes = cont_type_xml.css "Override"

            default_nodes.each do |node|
              # checking if extension type already present in the content types xml extracted from the first document
              if !default_extensions.include?(node["Extension"]) && !node.to_xml.empty?
                additional_cont_type_entries << node
                default_extensions << node["Extension"]    # extra extension type to be added to the content types XML
              end
            end

            override_nodes.each do |node|
              # checking if override content type info already present in the content types xml extracted from the first document
              if !override_partnames.include?(node["PartName"]) && !node.to_xml.empty?
                additional_cont_type_entries << node
                override_partnames << node["Partname"]       # extra content type info to be added to the content types XML
              end
            end
          end
        end

        zip_file.entries.each do |e|
          unless e.name == document_file_path || [RELATIONSHIP_FILE_PATH, CONTENT_TYPES_FILE, STYLES_FILE_PATH].include?(e.name)
            if e.name.include?("word/media/image")
              #  media files from header & footer from first document shouldn't be changed
              if head_foot_media["doc#{doc_cnt}"].include?(e.name.gsub("word/media/", ""))
                e_name = e.name
              else
                e_name = e.name#.gsub(/image[0-9]*./, "image#{cnt}.")
                # storing the old media file name to new media file name to mapping in the media hash
                media_hash["doc#{doc_cnt}"][e.name.gsub("word/media/", "")] = cnt
                cnt += 1
              end
              zos.put_next_entry(e_name)
              zos.print e.get_input_stream.read
            else
              # writing the files not needed to be edited back to the new zip (only from the first document, so as to avoid duplication)
              if doc_cnt == 0
                zos.put_next_entry(e.name)
                zos.print e.get_input_stream.read
              end
            end
          end
        end

        # updating the stlye ids in the table elements present in the document content XML
        doc_content = doc_cnt == 0 ? @main_body : Nokogiri::XML(zip_file.read(document_file_path))
        doc_content.xpath("//w:tbl").each do |tbl_node|
          style_last = tbl_node.xpath('.//w:tblStyle').last
          unless style_last.nil?
            val_attr = style_last.attributes['val']
            table_hash["doc#{doc_cnt}"][val_attr.value.to_s] = tbl_cnt
            val_attr.value = val_attr.value.gsub(/[0-9]+/, tbl_cnt.to_s)
            tbl_cnt += 1
          end
        end

        zip_file.entries.each do |e|
          # updating the relationship ids with the new media file names in the relationships XML
          if e.name == RELATIONSHIP_FILE_PATH
            rel_xml = doc_cnt == 0 ? @rel_doc : Nokogiri::XML(e.get_input_stream.read)

            rel_xml.css("Relationship").each do |node|
              next unless node.values.to_s.include?("image")

              i = media_hash["doc#{doc_cnt}"][node['Target'].to_s.gsub("media/", "")]
              target_val = node["Target"].gsub(/image[0-9]*./, "image#{i}.")
              rid_hash["doc#{doc_cnt}"][node['Id'].to_s] = i.to_s

              id_attr = node.attributes["Id"]
              new_id = id_attr.value.gsub(/[0-9]+/, i.to_s)
              if doc_cnt == 0
                node["Target"] = target_val
                id_attr.value = new_id
              else
                #  adding the extra relationship nodes for the media files to the relationship XML
                new_rel_node = "<Relationship Id=#{new_id} Type=#{node["Type"]} Target=#{target_val} />"
                @rel_doc.at('Relationships').add_child(new_rel_node)
              end
            end
          end

          # adding the table style information to the styles xml, if any tables present in the document being merged
          if e.name == STYLES_FILE_PATH
            style_xml = doc_cnt == 0 ? @style_doc : Nokogiri::XML(e.get_input_stream.read)
            table_nodes = style_xml.xpath('//w:style').select{ |n| n.attributes["type"].value == "table" }
            table_nodes = table_nodes.select{ |n| n.attributes["styleId"].value != "TableNormal" } if doc_cnt != 0

            table_nodes.each do |table_node|
              style_id_attr = table_node.attributes['styleId']
              tab_val = table_hash["doc#{doc_cnt}"][style_id_attr.value.to_s]
              style_id_attr.value = style_id_attr.value.gsub(/[0-9]+/, tab_val.to_s)

              # adding extra table style nodes to the styles xml, if any tables present in the document being merged
              @style_doc.xpath("//w:styles").children.last.add_next_sibling(table_node.to_xml) if doc_cnt != 0
            end
          end
        end

        # updting the id and rid values for every drawing element in the document XML with the new counters
        doc_content.xpath("//w:drawing").each do |dr_node|
          docPr_node = dr_node.xpath(".//wp:docPr").last
          docPr_node['id'] = doc_pr_id.to_s
          doc_pr_id += 1

          blip_node = dr_node.xpath(".//a:blip", NAMESPACES).last
          #  not all <w:drawing> are images and only image has <a:blip>
          next if blip_node.nil?
          embed_attr = blip_node.attributes["embed"]
          i = rid_hash["doc#{doc_cnt}"][embed_attr.value]
          embed_attr.value = embed_attr.value.gsub(/[0-9]+/, i)
        end

        if doc_cnt > 0
          # pulling out the <w:sectPr> element from the document body to be appended to the main document's body
          body_nodes = doc_content.xpath('//w:body').children[0..-2]

          # appending the body_nodes to main document's body
          @place.parent.children = body_nodes
        end

        doc_cnt += 1
      end

      # writing the updated styles XML to the new zip
      zos.put_next_entry(STYLES_FILE_PATH)
      zos.print @style_doc.to_xml

      # writing the updated relationships XML to the new zip
      zos.put_next_entry(RELATIONSHIP_FILE_PATH)
      zos.print @rel_doc.to_xml

      zos.put_next_entry(CONTENT_TYPES_FILE)
      additional_cont_type_entries.each do |node|
        # adding addtional content type nodes to the content type XML
        @cont_type_doc.at("Types").add_child(node)
      end
      # writing the updated content types XML to the new zip
      zos.print @cont_type_doc.to_xml

      # writing the updated document content XML to the new zip
      zos.put_next_entry(document_file_path)
      zos.print @main_document_xml.to_xml
    end

    # moving the temporary docx file to the final_path specified by the user
    FileUtils.mv(temp_file.path, final_path)
  end
end
