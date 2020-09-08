from tkinter import filedialog
from pydub import AudioSegment
from pathlib import Path
import os
import zipfile
import time


# Gets file path of user's selected file, filters for pptx.
filepath = filedialog.askopenfilename(initialdir = "~/",title = "Select file",filetypes = (("pptx files","*.pptx"),("all files","*.*")))

start_time = time.time()

filename = os.path.basename(filepath)[:-5]
directory = os.path.dirname(filepath)
os.chdir(directory)


# Renames file from .pptx extension to .zip.
os.rename(filepath, filepath[:-4] + 'zip')
zip_filepath = filepath[:-4] + 'zip'


# Unzips the newly-made zip file.
with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
    zip_ref.extractall(directory + '/arth_temp_files')


# Renames the .zip file back to a .pptx so it can be used while listening to the presentation.
os.rename(zip_filepath, filepath[:-4] + 'pptx')


# Navigate to the pptx audio files
os.chdir(directory + '/arth_temp_files/ppt/media')
MEDIA = Path(os.getcwd())


# Convert all m4a files to mp3 files.
os.system('for foo in *.m4a; do ffmpeg -i "$foo" -acodec libmp3lame -aq 2 "${foo%.m4a}.mp3"; done')


# Count how many audio files there are - wish there was a simpler way to do this.
count = 0
for audio in MEDIA.glob('*.mp3'):
    count += 1


# Combine all audio files (in order), then export
combined = AudioSegment.empty()
for i in range(count):
    sound_path = Path(os.getcwd() + '/media' + str(i+1) + '.mp3')
    combined += (AudioSegment.from_file(sound_path, "mp3"))
combined.export("../../../" + filename + ".mp3", format="mp3")


# Go back to original working directory, remove temp files (cleanup)
os.chdir(directory)
os.system('rm -rf arth_temp_files')

# Print timings
print("\n" + str(count) + "audio files successfully combined in %.2f seconds" % (time.time() - start_time))









