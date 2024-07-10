# DocumentGeneration
# Overview
This project demonstrates a simple ASP.NET MVC application for generating documents using customer data. The application allows users to fill out a form with customer details, save these details to a database, and generate a Word document based on a template using Syncfusion's DocIO library.

# Features
Form Submission: A user-friendly form to capture customer details.

Database Integration: Save and retrieve customer details from a database.

Document Generation: Generate a Word document using a predefined template with customer details.

# Technologies Used
  1.ASP.NET MVC
  
  2.ADO.NET for database interaction
  
  3.Syncfusion.DocIO for document generation
  
  4.SQL Server for database
  
# Prerequisites

Visual Studio 2019 or later

SQL Server 2017 or later

Syncfusion.DocIO (can be installed via NuGet)

# Folder Structure

Controllers: Contains the HomeController which handles form submission, document creation, and database operations.

Models: Contains the Customer model representing customer data and the DemoEntities DbContext for database access.

Views
Home
CreateDocument.cshtml: Form to capture customer details.
Index.cshtml: Default landing page.

Data: Folder containing the Word document template (SampleDocument.docx).

Getting Started
Clone the repository

bash
Copy code
git clone https://github.com/abhishekalandikar/DocumentGeneration.git
Open the project in Visual Studio

Configure the database

Ensure your SQL Server is running.
Update the connection string in Web.config to match your database settings.
Run the application

Press F5 to run the application.
Navigate to the CreateDocument page to fill out the form and generate a document.
