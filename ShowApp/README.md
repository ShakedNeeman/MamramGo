# Windows Service Display

This C# project demonstrates how to list Windows services using three different approaches:

1. **Windows Management Instrumentation (WMI)**
2. **Windows API (WinAPI)**
3. **Windows Registry**

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Classes](#classes)
- [Additional_Information](#additional_Information)

## Features

- List all running and installed Windows services
- Display services in a user-friendly format
- Use three different methods for fetching the service data

## Installation

1. Clone the repository: 
    ```bash
    git clone https://github.com/ShakedNeeman/MamramGo.git
    ```
2. Open the solution in Visual Studio.
3. Build and run the project.

## Usage

After running the application, a menu will be displayed. Users can select one of the options to display the services using the corresponding method. Each method may show varying levels of detail for each service.

To exit the application, select the `Exit` option from the menu.

### Menu Options

- **1: Using Registry**
- **2: Using WMI**
- **3: Using WinAPI**
- **0: Exit**

## Classes

### `Program`

This is the entry point of the application. It provides a menu for the user to choose the method for displaying services.

### `RegistryServiceDisplay`

This class fetches and displays the services using the Windows Registry. 

- `Display()`: Method to get the services and display their details.

### `WMIServiceDisplay`

This class fetches and displays the services using Windows Management Instrumentation (WMI).

- `Display()`: Method to get the services and display their details.

### `WinAPIServiceDisplay`

This class fetches and displays the services using the Windows API.

- `Display()`: Method to get the services and display their details.

## Additional_Information

The data displayed for each service might include:

- Service Name
- Display Name
- Service Type
- Start Type
- And more...