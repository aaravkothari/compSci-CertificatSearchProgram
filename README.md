# Certificate Search Program

## Overview

The Certificate Search Program provides users with a search interface for computer science-related certifications. Users can find suitable certificates from the `Certification_Data.xlsx` database by applying four different search parameters.

### Features

- **Industry Selection**: Choose a specific industry within Computer Science (e.g., Cybersecurity).
- **Organization Selection**: Select the certification provider (e.g., Microsoft) or opt for "Any" organization.
- **Cost Filter**: Set a maximum cost for the certification, with options to select from a dropdown or enter a custom value.
- **Location Preference**: Choose between online, in-person, or either option for certificate delivery.

After entering the search parameters, users can click the search button to view the matching certificates, including their names, costs, locations, and registration links. If no matches are found, users can reset and modify their search criteria.

## Installation

To run the program, ensure you have the following dependencies installed:

```bash
pip install customtkinter openpyxl
```

## Libraries Used

1. **customtkinter**: A custom Tkinter library for enhanced GUI components.
   - **Author**: Tom Schimansky
   - **Version**: 5.2.2
   - **Source**: [CustomTkinter GitHub Repository](https://github.com/TomSchimansky/CustomTkinter)

2. **openpyxl**: A library to read and write Excel files.
   - **Authors**: Eric Gazoni, Charlie Clark, and contributors
   - **Version**: 3.0.9
   - **Source**: [openpyxl Documentation](https://openpyxl.readthedocs.io/en/stable/)

3. **webbrowser**: A Python standard library for opening URLs in the web browser.
   - **Source**: [webbrowser Documentation](https://docs.python.org/3/library/webbrowser.html)

## Usage

1. **Load Certification Data**: The program loads data from the `Certification_Data.xlsx` file containing certification information.
2. **User Input**: Users select their desired parameters using dropdowns and a combobox.
3. **Search Execution**: On clicking the search button, the program matches the user’s input against the certification data.
4. **Results Display**: Matched certificates are displayed with their names, costs, locations, and links. Users can also reset their search parameters.

## Code Snippet

Here's a brief code snippet illustrating the search functionality:

```python
def searchCertificate(selection):
    # Converts cost keywords (from combobox dropdown) to numerical values
    if selection[2] == "Free":
        selection[2] = "0"
    if selection[2] == "Any":
        selection[2] = "2000"

    try:
        selection[2] = float(selection[2])
    except:
        searchOutput("Error")
        return

    matched = []
    # Matching logic here...
```

## Contributing

Contributions are welcome! If you have suggestions or improvements, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Personal Use

Feel free to use personal database, and fork and customize according to your preferences!