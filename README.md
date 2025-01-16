# Autocell: Advanced Automation Tool

#### Video Demo: [Watch Here](<URL HERE>)

#### Description **Autocell** is a sophisticated Python-based automation application designed to handle repetitive and time-consuming tasks efficiently. It is tailored for professionals and individuals who seek to enhance their productivity by automating workflows in Excel, Word, and desktop environments while ensuring secure user authentication. Built with PyQt5, the application features an intuitive graphical user interface (GUI) that provides a seamless experience for users of all skill levels. By leveraging powerful Python libraries like `pandas`, `openpyxl`, and `cryptography`, Autocell delivers a robust, secure, and user-friendly solution to streamline operations. One of Autocell’s standout features is its **Excel Automation** module. This tool allows users to import and export data from various formats such as CSV and TXT files. The data cleaning feature ensures that duplicate entries are removed, and datasets are standardized for further analysis. For users who require unit conversions within their datasets, Autocell provides a straightforward way to transform column data using user-defined conversion factors. Additionally, the application includes a financial reporting tool that generates visual financial reports, complete with charts, enabling users to present data insights effectively. The **Word Automation** module simplifies the creation and manipulation of Word documents. Whether users need to create tables from scratch, import data from text files, or convert CSV files into professional Word documents, Autocell offers comprehensive solutions. The auto table maker supports both manual input and automated generation from text files, providing flexibility for diverse use cases. The application also ensures that converted documents maintain consistent formatting, saving users significant time compared to manual adjustments.Autocell extends its capabilities to desktop automation with tools like the **Auto Clicker** and **Keyboard Automation**. The auto clicker allows users to automate left, right, and double-click actions with customizable delays, making it ideal for repetitive desktop tasks. For text-heavy workflows, the keyboard automation tool can simulate keystrokes, typing out the contents of `.txt` files at user-defined speeds. These features are particularly useful for tasks requiring constant interaction with a computer, reducing physical strain and increasing efficiency. To ensure data security, Autocell incorporates a robust **User Authentication** system. Users can register and log in with secure credentials, which are validated against a strict password policy requiring a mix of uppercase and lowercase letters, numbers, and special characters. The application stores encrypted user credentials in a MySQL database, using the `cryptography` library to safeguard sensitive information. This ensures that user data is protected against unauthorized access, making Autocell a reliable tool for secure operations. The application’s development reflects deliberate design choices aimed at maximizing usability and functionality. PyQt5 was chosen as the GUI framework for its cross-platform compatibility and extensive widget library, which allowed for a sleek and responsive interface. For database management, MySQL was selected due to its scalability and ease of integration with Python. The encryption layer provided by the `cryptography` library ensures that user data remains secure, aligning with modern security standards. Libraries like `pandas` and `openpyxl` were integrated for their robust data manipulation capabilities, enabling efficient processing of large datasets. Autocell’s project structure is designed to keep functionalities modular and organized. The `app.py` file serves as the main entry point, initializing the GUI and orchestrating interactions between different modules. The `database.py` file manages all authentication-related operations, including user registration, login, and password encryption. The `excel_automation.py` module handles Excel-specific tasks such as data cleaning and report generation, while the `word_automation.py` module is dedicated to Word document processing. The `autoclicker.py` file contains the logic for mouse and keyboard automation, completing the desktop automation suite. A `requirements.txt` file lists all necessary dependencies, ensuring that the application can be set up seamlessly. Installing and using Autocell is straightforward. After cloning the repository, users can install the required dependencies using `pip` and set up the MySQL database for user authentication. The application is then launched via the `app.py` file, presenting users with a clean and intuitive interface. From there, users can access various modules, select tasks, and follow on-screen prompts to complete their workflows. Detailed instructions ensure that even users with minimal technical experience can take full advantage of Autocell’s capabilities. Autocell’s development journey was not without challenges. Ensuring cross-platform compatibility required extensive testing to address differences in system environments. Implementing robust security features while maintaining ease of use demanded a careful balance between complexity and accessibility. Designing an intuitive GUI required iterative feedback and refinements to ensure that the application met user expectations. These challenges, however, shaped Autocell into a well-rounded tool that combines functionality, security, and usability. Looking forward, Autocell has significant potential for future enhancements. Integrating cloud storage options would allow users to save and access files remotely, making the application more versatile. Adding advanced analytics tools for Excel automation could provide users with deeper insights into their data. Customizable GUI themes would enable users to personalize their experience, further enhancing usability. These planned improvements demonstrate a commitment to continuous development and user satisfaction. In conclusion, **Autocell** is more than just an automation tool—it is a comprehensive solution for modern productivity challenges. By addressing repetitive workflows with advanced automation features and prioritizing data security, Autocell empowers users to focus on what matters most. Whether you are a professional managing large datasets or an individual seeking to simplify desktop tasks, Autocell is a powerful ally in achieving efficiency and excellence.

### Key Features

1. **Excel Automation**:
   - Import/export data from CSV and TXT files.
   - Clean and remove duplicate entries within datasets.
   - Convert units in specified columns.
   - Generate comprehensive visual financial reports with charts.

2. **Word Automation**:
   - Create tables manually or from text files.
   - Convert plain text files or CSV files into formatted Word documents.

3. **Desktop Task Automation**:
   - Automate mouse clicks (single, double, right) with customizable delays.
   - Simulate keystrokes by typing text from a `.txt` file.

4. **Secure User Authentication**:
   - MySQL integration for storing encrypted user credentials.
   - Password encryption using `cryptography` ensures security.

### Project Files

- **`app.py`**: The main application file initializing the GUI and integrating all features.
- **`database.py`**: Handles user authentication, registration, and secure password encryption.
- **`excel_automation.py`**: Implements Excel-related functionalities such as data cleaning and financial report generation.
- **`word_automation.py`**: Contains logic for Word document creation and manipulation.
- **`autoclicker.py`**: Enables desktop automation features like mouse clicks and keyboard typing.
- **`requirements.txt`**: A list of all the Python dependencies required to run the project.
- **`README.md`**: Comprehensive documentation for the project, explaining its functionality and setup.

### Design Choices

- **GUI Framework**: PyQt5 was chosen for its robust widget system and cross-platform compatibility, allowing for a professional and interactive user experience.
- **Database**: MySQL was selected for its reliability and support for secure data storage, ensuring robust authentication.
- **Encryption**: The `cryptography` library was used to encrypt passwords, adding a layer of security.
- **File Handling**: `pandas` and `openpyxl` provide powerful tools for processing and manipulating data in Excel files.

### Installation Instructions

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/your-username/autocell.git
   cd autocell
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set Up MySQL Database**:
   - Install MySQL Server if not already installed.
   - Create a database and configure connection details in `database.py`.

4. **Run the Application**:
   ```bash
   python app.py
   ```

### Usage Instructions

- **Authentication**: Register or log in with a secure username and password.
- **Excel Automation**: Choose a feature like data cleaning or report generation, provide the input file, and follow on-screen prompts.
- **Word Automation**: Select a feature to create a table or convert files to Word format.
- **Auto Clicker**: Set up the type of mouse click and delay, then press `q` to stop the automation.
- **Keyboard Automation**: Load a `.txt` file, specify typing speed, and observe the simulated keystrokes.

### Challenges and Future Enhancements

- **Challenges**:
  - Ensuring cross-platform compatibility required extensive testing.
  - Balancing security and usability in the authentication system.
  - Designing an intuitive GUI with PyQt5.

- **Future Enhancements**:
  - Adding cloud storage options for file management.
  - Integrating advanced analytics for Excel automation.
  - Allowing users to customize the GUI theme.

### Conclusion

Autocell is a powerful and versatile tool that simplifies tedious tasks, enhances productivity, and ensures security. With its user-friendly interface and robust functionality, it is an invaluable resource for professionals and individuals alike.

