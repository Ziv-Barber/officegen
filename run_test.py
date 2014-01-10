from subprocess import call
import shutil
import subprocess

call("npm test", shell=True)

#shutil.rmtree( "out", True )
#call(["./tools/unzip.exe", "out.pptx", "-d", "out"])

call('start out.pptx', shell=True)
