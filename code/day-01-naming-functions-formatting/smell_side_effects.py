def check_password(user, password):
    if user.password_hash == hash(password):
        Session.initialize(user)  # side effect - the name never mentions this
        return True
    return False

# reads like a pure question: "is this the right password?"
# a caller who just wants to confirm a password before, say,
# showing a "delete account?" dialog just started a session
# as an invisible side effect
