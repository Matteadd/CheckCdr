import sys, os
from pprint import pprint
try:
    print(1/0)
except Exception as e:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    print(sys.exc_info())
    print(pprint(exc_tb))
    # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
print(f"Error in Tool.\n"+
      f"Error type: {exc_type}\n"+
      f"Error line: {exc_tb.tb_lineno}")
