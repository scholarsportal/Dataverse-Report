import sys
import time
from metrics.excel import metrics

start_time = time.time()

if sys.version_info[0] < 3:
    print(sys.version_info)
    sys.exit("Python 3.x version is required to run this program.")

metrics.Metrics.generate()

mins, secs = divmod(time.time() - start_time, 60)
sys.stdout.write("Total time taken: %d minutes %d seconds.\n" % (mins, secs))
