import os
import array as arr 
import ffmpeg
from moviepy.editor import *
from natsort import natsorted

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
			#print("STOP")
			
		else:
			recursiveFind(path, master_array) #keeps going inside files
			#print("KEEP GOING")

	return






#concatinating vidoes
def merge(master_array):

	#print(master_array)

	clips = []

	for paths in master_array:
		if paths.endswith(".mp4"):
			#print(paths) 

			filePath = os.paths.join(root, master_array)
			video = VideoFileClip(filePath)
			clips.append(video)
			#print(video)

	final_clip = concatenate_videoclips(clips)
	#final_clip.to_videofile("output.mp4", fps = 24, remove_temp = False)


	#printf "file '%s'\n" *.mp4 > list.txt #create list of all mp4s in dc

	return

"""
#concatinating vidoes
def merge(master_array):

	video_clips = 0
	final_clip = 0

	video_arr = []

	for path in master_array:
	   file_path = os.path.join(root, master_array)
      #video = VideoFileClip(file_path) #make array of this and pass to final clip

   	video_arr.append(VideoFileClip(file_path))

	final_clip = concatenate_videoclips(video_arr)
	final_clip.to_videofile("output.mp4", fps=24, remove_temp=False)

	print(final_clip)

	return

"""




#Main function that calls recursive function
def main():
	print('Current working directory : ', os.getcwd())
	folder1_path = os.getcwd() + '/dates'
	print(os.listdir(os.getcwd()))

	master_array = [] #fill array with contents of directory

	recursiveFind(folder1_path, master_array)

	merge(master_array)

	#print(master_array)

	return

main() 


