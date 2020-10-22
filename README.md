# GUI-to-communicate-with-embedded-systems
Embedded systems lab project at TU Chemnitz

# About the project
Welcome to the application CIFEA. CIFEA stands for "Computer Interface For Embedded Applications". This project has been developed as part of "Project lab embedded systems(PLES)" course in the "Measurement and sensor technology (MST)" department at TU Chemnitz. The application mainly has three featues. First feature is to set different parameters for different electro-chemistry methods and send them to a microcontroller over serial port using UART (Universal asynchronous receiver transmitter) communication protocol. Second feature is to receive measurement data from microcontroller over serial port and draw a real time plot of the data. Third feature is to export the received data to an excel sheet on user request for further analysis of the measurement data.

# About the folder CIFEA
This folder contains the microsoft visual studio project file for GUI of the project.
Redirect to "CIFEA/bin/Debug/CIFEA.exe" and download the CIFEA.exe file and run to experience the GUI of the project.

# About the folder CIFEA_Controller
This folder contains the STM32CubeIDE project file for controller to receive and send data over UART protocol.
Redirect to CIFEA_Controller/Core/Src/main.c to check the controller coding part.
