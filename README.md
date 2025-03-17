# Word Format Checker Add-in  

A simple Microsoft Word Add-in that checks text formatting (bold, underline, and font size) using the **Word JavaScript API**.

## Features  
- Detects if the **first word** is bold.  
- Checks if the **second word** is underlined.  
- Reads the **font size** of the third word.  

## Installation  
1. Clone the repository:  
   ```sh
   git clone https://github.com/ScottTan98/Word-Format-Checker-Add-in-Office-.git
   cd word_format_checker
   ```
2. Install dependencies:  
   ```sh
   npm install
   ```
3. Run the add-in:  
   ```sh
   npm start
   ```

## Usage  
1. Open **Microsoft Word**.  
2. Sideload the add-in (refer to [Microsoft Docs](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing)).  
3. Click the **"Run"** button in the task pane to check the text formatting.  

## Testing  
Run unit tests with:  
```sh
npm test
```

## Technologies  
- **Office.js** (Word JavaScript API)  
- **React** (if applicable)  
- **Jest** (for testing)  

## License  
MIT License  