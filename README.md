# TimeTable-Generator
Class Timetable Generator
    Work in Progress !
If you'd like to contribute:
Feel free to fork the repository
Create your feature branch
Implement your changes
Submit a pull request
Contact chisomobanja@gmail.com with any questions

A Java-based command-line application for creating and managing class timetables with Excel export functionality.
Features

 Add, edit, and delete subjects
 Automatic time conflict detection
 Multiple sorting options (by day/time, name, or duration)
 Excel export with two views:

Weekly schedule view
List view of all subjects


 Flexible time slot management
 Input validation
 Subject duration tracking

Technical Requirements

Java 22
Maven 3.6+
Dependencies:

Apache POI 5.2.3 (for Excel operations)



Installation

Clone the repository:

git clone https://github.com/chisomobanja/timetable-generator.git
cd timetable-generator

Build the project:

mvn clean install

Run the application:

java -jar target/timetable-application-1.0-SNAPSHOT.jar
Usage
The application provides a command-line interface with the following options:

Add Subject

Enter subject name
Specify duration in minutes
Choose day (1-7, where 1 is Monday)
Set start time (HH:mm format)


Edit Subject

Select subject by ID
Modify any attribute
Skip unchanged fields by pressing Enter


Delete Subject

Remove subjects by ID


View All Subjects

Display formatted table of all subjects


Check Time Conflicts

Automatically detect overlapping schedules


Sort Subjects

By day and time
By subject name
By duration


Export to Excel

Creates 'Timetable.xlsx' with two sheets:

Weekly view (visual timetable)
List view (detailed subject information)





Known Issues

The weekly view in Excel might not properly merge cells for subjects longer than 30 minutes
Time conflict warning allows override without additional validation
No data persistence between sessions


Contributing
This project is open for improvements! Areas that need attention:

Data persistence (database/file storage)
GUI implementation
Better Excel formatting for the weekly view
Additional validation rules
Unit tests implementation



Building from Source
mvn clean package
This will create an executable JAR in the target directory.
Technical Details

Uses Java's LocalTime for time management
Implements Comparable interface for subject sorting
Apache POI for Excel file generation
Maven for dependency management and building

License
MIT License - Feel free to use and modify the code
Contact

Email: [chisomobanja@gmail.com]
GitHub: [chisomobanja]

Acknowledgments

Apache POI library for Excel export functionality
Maven build system
