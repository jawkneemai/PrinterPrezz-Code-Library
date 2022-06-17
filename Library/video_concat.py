import os
import array as arr 
import ffmpeg
from moviepy.editor import *
from natsort import natsorted
import numpy as np

#Recursice Function that finds path to mp4 vidoes
#Stores path in temp_array which is sorted
def recursiveFind(folder_path, master_array):

	temp_array = os.listdir(folder_path)

	temp_array.sort() #sorts videos
	#print(temp_array)

	for path in temp_array: 
		path = folder_path + '/' + path
		if ".DS_Store" in path:

			continue

		if ".mp4" in path: #if reaches mp4 videos
			master_array.append(path)
			#print(path)
			#print("STOP")
			
		else:
			recursiveFind(path, master_array) #keeps going inside files
			#print("KEEP GOING")

	#print(split_array) #print(split_array[0]) and so on

	return 





#concatinating vidoes
def merge(split_array, master_array):

	print(len(master_array))
	i = 0

	size = len(master_array)/50
	print("Size: ", size)

	split_array =  np.array_split(master_array, size) 

	while i < size:

		clips = []

		duration = 0

		print(len(split_array[i]))

		for paths in split_array[i]:
			#print(type(paths))
			if paths.endswith(".mp4"):
				print(paths)

				video = VideoFileClip(paths)

				print(video.fps)

				duration = duration + video.duration
				
				clips.append(video)
				
				#print(video)
		print(i)
		#print(duration)
		i = i + 1

		#print(clips)

		final_clip = concatenate_videoclips(clips)

		#final_clip = (final_clip.subclip(t_end=round(final_clip.duration - 1.0/final_clip.fps), 2))

		final_clip.duration = round(final_clip.duration, 2)

		#if final_clip != 0:
		#print(final_clip.duration)
		final_clip.write_videofile("output.mp4", fps=15.0, remove_temp=False)

		for video in clips:
			video.close()

	return split_array





#Main function that calls recursive function
def main():
	print('Current working directory : ', os.getcwd())
	folder1_path = os.getcwd() + '/data/dates'
	print(os.listdir(os.getcwd()))

	master_array = [] #fill array with contents of directory

	split_array = recursiveFind(folder1_path, master_array)

	#recursiveFind(folder1_path, master_array)

	#print(split_array)

	merge(split_array, master_array)

	#print(master_array)

	return

main() 


