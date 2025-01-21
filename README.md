# Excel to CSV Converter

A web-based application for converting Excel files to a specific CSV format for Ulverscroft data processing.

## Features

### File Upload
- Accepts Excel files (.xlsx, .xls)
- Processes book data including ISBN, titles, dimensions, and page counts
- Automatically validates and transforms data according to business rules

### Data Preview
- Interactive table showing converted data before download
- Displays key fields:
  - ISBN
  - Title
  - Height
  - Width
  - Pages
  - Spine Size
- Row-by-row deletion capability
- Clear All functionality to reset the application

### Data Transformation Rules

#### Dimension Adjustments
- Width Transformation:
  - If width is 155, automatically converts to 156
  - All other width values remain unchanged
  - Default value: 156
- Height Transformation:
  - If height is 233, automatically converts to 234
  - All other height values remain unchanged
  - Default value: 234

### CSV Output
- Generates CSV files without headers
- Filename Format: `ulverscroft_itemTemplate[YYYY]_[MM]_[DD]_[HH]_[mm]_[ss]_[ms].csv`
- Example: `ulverscroft_itemTemplate2025_01_21_15_30_45_123.csv`

### Template Structure
The output CSV includes the following columns:
1. ISBN
2. NEW
3. 9781399183413 (ISBN reference)
4. The Death of Dora Black (Title)
5. Limp
6. Gloss
7. 234 (Height)
8. 156 (Width)
9. 19 (Spine Size)
10. 342 (Page Extent)
11. 60
12. Holmen Book Cream 60 gsm
13. N

Default values are automatically applied for standard fields.

## Usage Instructions

1. **Upload File**
   - Click the file input or drag and drop an Excel file
   - The system will automatically process the file upon selection

2. **Review Data**
   - Examine the converted data in the preview table
   - Verify that dimensions have been correctly transformed
   - Check that all required fields are present

3. **Manage Rows**
   - Remove unwanted rows using the delete button (trash icon)
   - Use "Clear All" to start over if needed

4. **Download CSV**
   - Click "Download CSV" to generate the final file
   - The file will be saved with the correct naming convention
   - Verify the downloaded file contains all required data

## Technical Details

### Technologies Used
- HTML5
- Bootstrap 5.3.2 for UI components
- SheetJS for Excel file processing
- PapaParse for CSV generation
- Font Awesome 6.4.0 for icons

### Dependencies
- bootstrap.min.css
- xlsx.full.min.js
- papaparse.min.js
- all.min.css (Font Awesome)

### Browser Compatibility
- Chrome (recommended)
- Firefox
- Safari
- Edge

## Error Handling
- Provides clear status messages for all operations
- Validates file input type
- Handles missing or invalid data gracefully
- Shows error messages when processing fails

## Data Validation
- Ensures correct data type conversion
- Applies business rules consistently
- Maintains data integrity throughout the conversion process