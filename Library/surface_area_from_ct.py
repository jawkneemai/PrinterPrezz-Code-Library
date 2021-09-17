# Johnathan Mai
# Python 3
# Requires Pillow (python imaging module), type in cmd prompt:
# python -m pip install --upgrade Pillow
# and scikit-image:
# python -m pip install -U scikit-image

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
import matplotlib.pyplot as plt
from PIL import Image
from skimage.filters import sobel
from skimage import segmentation
from skimage.color import label2rgb
from scipy import ndimage as ndi




def contour_image(image):

	# Get Image matrix, data
	imarray = np.array(image)

	# Region-based edge detection to get contours
	elevation_map = sobel(imarray)

	markers = np.zeros_like(imarray)
	markers[imarray < 30] = 1
	markers[imarray > 150] = 2

	seg = segmentation.watershed(elevation_map, markers)
	seg = ndi.binary_fill_holes(seg-1)
	labeled, _ = ndi.label(seg)
	image_label_overlay = label2rgb(labeled, image=imarray, bg_label=0)

	fig, ax = plt.subplots(figsize=(4,3))
	ax.imshow(imarray, cmap=plt.cm.gray)
	contours = ax.contour(seg, [1], linewidths=1.2, colors='y')
	plt.close('all')

	return contours




def calculate_contour_length(contour, voxel_size):
	old_x = -1
	old_y = -1
	length = 0

	for coordinates in contour: # At individual contour
		if old_x == -1:
			old_x = coordinates[0]
			old_y = coordinates[1]
		else:
			temp_x = coordinates[0]
			temp_y = coordinates[1]
			line_distance = sqrt((temp_x - old_x)**2 + (temp_y - old_y)**2)
			actual_distance = voxel_size * line_distance
			length = length + actual_distance
			old_x = coordinates[0]
			old_y = coordinates[1]

	return length




def main():
	root = tkint.Tk()
	root.withdraw()
	file_path = filedialog.askopenfilename()
	im = Image.open(file_path)
	voxel_size = 0.074133558


	total_sa = 0

	for i in range(1, im.n_frames):
		print(i)
		im.seek(i)
		contours = contour_image(im)

		# Calculating length of contours
		if contours.allsegs[0]:
			print('theres contours')
			contour_lengths = []
			for contour in contours.allsegs[0]: # Set of all contours made on image
				contour_lengths.append(calculate_contour_length(contour, voxel_size))
				
			temp_sa = sum(contour_lengths) * voxel_size
			total_sa = total_sa + temp_sa
			print(temp_sa)
			print(total_sa)
		else:
			print('no contours')

	#plt.show()

	return



if __name__ == '__main__':
	main()

print('Imported surface_area_from_ct.py!')