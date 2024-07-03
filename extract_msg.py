import extract_msg


def msg_to_mht(msg_file, mht_file):
    msg = extract_msg.Message(msg_file)
    msg.save()


msg_to_mht("example.msg", "example.mht")
