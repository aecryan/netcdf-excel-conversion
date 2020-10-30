# netcdf-excel-conversion
This repository contains code used to convert between files in netCDF (.nc) to Excel (.xlsx, .xls) and from Excel to netCDF. In the Excel -> netCDF direction, it is set up to convert individual sheets within a workbook to netCDF format.

## How to use these scripts
This repository contains two scripts for conversion between Excel (.xlxs and .xls) and netCDF (.nc) files. They are called, ExcelToNetcdf.py and NetcdfToExcel.py, respectively.
You can run them on their own or in a Docker container. The Docker container can be built via the Dockerfile in the repository, or from the image on DockerHub https://hub.docker.com/r/acryan/netcdf-excel-conversion.

## Adding your files for conversion
1. After cloning the repository, on your local machine, create an empty folder called `data/` and two other folders inside it called `input/` and `output/`.
2. Copy all files to be converted into the `data/input/` folder. These files should be Excel and/or netCDF files. 

## Steps to run the code in the netcdf-excel-conversion Docker container
1. `cd` into the root directory of this repository on your local machine. **Remember:** You must have added the three folders described above and copied your files for conversion into the `data/input/` folder.
2. In the root directory, start the Docker container: `docker run -it -v $(pwd):/work acryan/netcdf-excel-conversion /bin/bash` (you will end up with a bash prompt inside the container)
3. Go to the directory that contains the python script: `cd /work`
4. Run the python script: `python ExcelToNetcdf.py`. Each time you convert a file, you will be prompted to input the name of the original file. In the case of converting Excel files, you will also need to input the name of the individual sheet to be converted.
5. To exit the container, run `exit`

## Where are my converted files?
All converted files will be written to the `data/output` folder you created in Step 1.
