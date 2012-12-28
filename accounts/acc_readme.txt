Account Files

   userid.txt    = User reference file (unused right now)
   userid_bl.txt = User buddy list
   reference.txt = All accounts, with passwords, seperated by =. i.e. UserID=UserPW

File Structures

userid_bl.txt

    All these files are simple INI files. The structure is quite straight forward.

	[Buddylist]
	Total=2 # Number of 'Buddy' catagories below. Each one is a buddy in their list.

	[Buddy_1]
	UserID=admin # Buddy's service ID.
	Title=Administrator # How the buddy's name appears in their list.

	[Buddy_2]
	UserID=you # Buddy's service ID.
	Title=Youralia # How the buddy's name appears in their list.

reference.txt

    Reference file used by server for quick authentication checking. User id followed by
    password.

	[Accounts]
	userid=userpass
