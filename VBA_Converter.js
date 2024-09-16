function convertToVbaString(input) {
    // Split the input by new lines
    let lines = input.split(/\r?\n/);
    
    // Initialize the VBA string
    let vbaString = 'Dim vbaCode As String\nvbaCode = ';

    // Loop through each line, escape quotes, and format for VBA
    lines.forEach((line, index) => {
        // Escape double quotes
        line = line.replace(/"/g, '""');

        // If it's the last line, don't add the & vbCrLf at the end
        if (index === lines.length - 1) {
            vbaString += '"' + line + '"';
        } else {
            vbaString += '"' + line + '" & vbCrLf & _\n';
        }
    });

    return vbaString;
}

// Example usage
const htmlCssJs = ``;

console.log(convertToVbaString(htmlCssJs));
