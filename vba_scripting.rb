class DocumentFormatter
  require 'win32ole'

  def run_saved_macro(file_path, macro_name)
    puts "File Path : #{file_path}"
    puts "Macro Name: #{macro_name}"
    
    begin
      check_file_present(file_path)

      word = WIN32OLE.new('word.application')
     # word.visible = true
      doc = word.Documents.Open(file_path)
      word.run macro_name
      doc.save
      puts "-------Macros Run Successfully--------"    
    rescue Exception => e
      puts "Error: #{e.message}"
    ensure
      word.Quit if word
    end
  end

  private
  def check_file_present(file_path)
   raise Exception.new("File #{file_path} Not Found") if !File.file?(file_path)
  end 
end 






file_path = ARGV[0]   # C:\Users\Administrator\Documents\VBA scripting program\sample1.doc
macro_name = ARGV[1] 	# sample_macro

DocumentFormatter.new().run_saved_macro(file_path, macro_name)