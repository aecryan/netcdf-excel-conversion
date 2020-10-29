# netcdf-excel-conversion
This repository contains code used to convert between files in NetCDF (.nc) to Excel (.xlsx) and from .xlsx to NetCDF (.nc). 

## How to use these scripts
This repository contains two scripts for conversion between Excel (.xlxs and .xls) and netCDF (.nc) files. They are called, ExcelToNetcdf.py and NetcdfToExcel.py, respectively.
You can run them on their own or in a Docker container. The Docker container can be built from the image on DockerHub https://hub.docker.com/r/acryan/netcdf-excel-conversion or by using step 2 below which accesses the Dockerfile in the repository.

## Converting your files
1. After cloning the repository, on your local machine, create an empty folder called `data/` and two other folders inside it called `input/` and `output/`.
2. Copy all files to be converted into the `data/input/` folder. These files should be Excel and/or netCDF files. 
3. Run the python files. Each time you convert a file, you will be prompted to input the name of the original file.
4. All converted files will be written to the `data/output` folder.


## Steps to run the code
1. After cloning this repository to your local machine, `cd` into that directory
2. Start the Docker container: `docker run -it -v $(pwd):/work netcdf-excel-conversion` /bin/bash (you will end up with a bash prompt inside the container)
3. Go to the directory that contains the python script: `cd /work`
4. Run the python script: `python ExcelToNetcdf.py`
5. To exit the container, run `exit`
