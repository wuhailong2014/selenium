from my_class import *

username, password = read_user_info("json.json")
get = OPERATOR()
get.get_cookies(username=username,  password=password)
del get

op = OPERATOR()
op.op()
while True:
    try:
        op.do_work()
    except Exception:
        time.sleep(1.2)
