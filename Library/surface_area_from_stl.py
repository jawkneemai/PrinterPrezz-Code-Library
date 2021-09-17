# Johnathan Mai
# Python 3
# Requires numpy-stl, run in console:
# pip install numpy-stl

'''
Directory of Functions



'''

# Python Imports
import tkinter as tkint
from tkinter import filedialog
from math import sqrt

# Non-Python Imports
from Library import ancillary
import numpy as np
import math
from stl import mesh as mesh
from matplotlib import pyplot
from mpl_toolkits import mplot3d






def main():
	root = tkint.Tk()
	root.withdraw()
	file_path = filedialog.askopenfilename()
	mesh1 = mesh.Mesh.from_file(file_path)

	# Plot STL
	'''
	figure = pyplot.figure()
	axes = mplot3d.Axes3D(figure)
	axes.add_collection3d(mplot3d.art3d.Poly3DCollection(mesh1.vectors))
	scale = mesh1.points.flatten()
	axes.auto_scale_xyz(scale, scale, scale)

	pyplot.show()
	'''

	print(mesh1.areas.sum())

	return



if __name__ == '__main__':
	main()

print('Imported surface_area_from_stl.py!')