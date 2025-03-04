require 'roo'
require 'csv'
require 'dotenv/load'

class Dishwasher
  def self.load_marketplace_mappings
    # Load marketplace mappings from CSV
    mappings = {}
    CSV.foreach(ENV['MARKETPLACES_CSV_PATH'], headers: true) do |row|
      mappings[row['name']] = row['id']
    end
    mappings
  end

  def self.update_businesses_from_excel(start_line = ENV['START_LINE'].to_i, end_line = ENV['END_LINE'].to_i)
    # Load marketplace mappings
    marketplace_mappings = load_marketplace_mappings

    # Load the Excel file
    xlsx = Roo::Spreadsheet.open(ENV['EXCEL_FILE_PATH'])

    # Get the first sheet
    sheet = xlsx.sheet(0)

    # Initialize a hash to group businesses by trading name
    corrections_by_name = {}

    # Process each row within the specified range
    (start_line..end_line).each do |row_num|
      row = sheet.row(row_num)
      next unless row[0] # Skip if row is empty

      # Map columns to their values
      id = row[0]
      trading_name = row[1]
      trading_name_corrected = row[2]
      is_omp = row[3]
      is_omp_corrected = row[4]
      company_number = row[5]
      company_number_corrected = row[6]
      legal_name = row[9]
      legal_name_corrected = row[10]
      marketplace_name = row[13] # Column N

      # Skip if no corrections needed
      next if (trading_name_corrected.nil? || trading_name_corrected.to_s.strip.empty?) &&
              (company_number_corrected.nil? || company_number_corrected.to_s.strip.empty?) &&
              (legal_name_corrected.nil? || legal_name_corrected.to_s.strip.empty?) &&
              (is_omp_corrected.nil? || is_omp_corrected.to_s.strip.empty?)

      # Get the trading name for grouping
      group_name = trading_name_corrected.to_s.strip if !trading_name_corrected.nil? && !trading_name_corrected.to_s.strip.empty?
      group_name ||= trading_name

      # Initialize or update the group
      corrections_by_name[group_name] ||= {
        ids: [],
        updates: {}
      }

      # Add the business ID to the group
      corrections_by_name[group_name][:ids] << id

      # Update the corrections with any non-nil values
      if !trading_name_corrected.nil? && !trading_name_corrected.to_s.strip.empty?
        corrections_by_name[group_name][:updates][:trading_name] = trading_name_corrected.to_s.strip
      end

      if !company_number_corrected.nil? && !company_number_corrected.to_s.strip.empty?
        # If company number is corrected to "0", set it to nil
        if company_number_corrected.to_s.strip == "0"
          corrections_by_name[group_name][:updates][:company_number] = nil
        else
          corrections_by_name[group_name][:updates][:company_number] = company_number_corrected.to_s.strip
        end
      end

      if !legal_name_corrected.nil? && !legal_name_corrected.to_s.strip.empty?
        corrections_by_name[group_name][:updates][:legal_name] = legal_name_corrected.to_s.strip
      end

      # Handle online marketplace updates
      if !is_omp_corrected.nil? && is_omp_corrected.to_s.strip.downcase == 'yes' && !marketplace_name.nil?
        marketplace_id = marketplace_mappings[marketplace_name.to_s.strip]
        if marketplace_id
          corrections_by_name[group_name][:updates][:online_marketplace_id] = marketplace_id
        end
      end
    end

    # Generate update commands
    output = []
    corrections_by_name.each do |_, correction_data|
      business_ids = correction_data[:ids]
      updates = correction_data[:updates].map { |field, value| "#{field}: '#{value}'" }.join(',')
      output << "business_ids = [#{business_ids.join(',')}]"
      output << "Business.where(id: business_ids).update(#{updates})"
      output << "\n"
    end

    # Write to file
    File.write(ENV['OUTPUT_FILE_PATH'], output.join("\n"))
    puts "Update commands have been written to #{ENV['OUTPUT_FILE_PATH']}"
  end
end

# Run the script with the specified range
Dishwasher.update_businesses_from_excel
