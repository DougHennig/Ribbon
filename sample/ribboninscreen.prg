lcCurrFolder = addbs(justpath(sys(16)))
set path to fullpath('..\', lcCurrFolder)

lcImagePath = sys(16)
lcImagePath = addbs(justpath(substr(lcImagePath, at(' ', lcImagePath, 2) + 1)))

* Add a ribbon to _SCREEN.

if type('_screen.oRibbon') <> 'O'
	_screen.NewObject('oRibbon', 'SFRibbon', 'SFRibbon.vcx')
	with _screen.oRibbon
		.Visible = .T.

* Display the RibbonDisplay button.

		.AllowShowTabsOnly = .T.

* Set up the Home tab.

		loTab = .AddTab('Home')
		with loTab
			.Caption = 'Home'

* Set up the New section.

			loSection = .AddSection()
			with loSection
				.Caption = 'New'
				loButton = .AddButton()
				with loButton
					.Caption = 'New' + chr(13) + 'Email'
					.Image   = lcImagePath + 'newemail.png'
					.Command = "messagebox('New email')"
				endwith
				loButton = .AddButton('NewItems')
				with loButton
					.Caption = 'New' + chr(13) + 'Items'
					.Image   = lcImagePath + 'newitems.png'
					.AddBar('E-\<mail Message', "messagebox('Email Message')", ;
						lcImagePath + 'newemailsmall.png')
					.AddBar('\<Appointment', "messagebox('Appointment')", ;
						lcImagePath + 'appointmentsmall.png')
					.AddBar('M\<eeting', "messagebox('Meeting')", ;
						lcImagePath + 'meetingsmall.png')
					.AddBar('Gro\<up', "messagebox('Group')", ;
						lcImagePath + 'groupsmall.png')
					.AddBar('\<Contact', "messagebox('Contact')", ;
						lcImagePath + 'contactsmall.png')
					.AddBar('\<Task', "messagebox('Task')", ;
						lcImagePath + 'tasksmall.png')
					.AddBar()
					loBar = .AddBar('E-mail Message \<Using')
					loBar.AddBar('\<More Stationery...', "messagebox('More stationery')")
					loBar.AddBar()
					loBar.AddBar('\<Plain Text')
					loBar.AddBar('\<Rich Text')
					loBar.AddBar('\<HTML')
					loBar = .AddBar('More \<Items', "messagebox('More Items')")
					loBar.AddBar('\<Post in This Folder')
					loBar.AddBar('Contact \<Group')
					loBar.AddBar('Task Re\<quest')
					.AddBar('Teams Meeti\<ng', "messagebox('Teams Meeting')", ;
						lcImagePath + 'teamsmeetingsmall.png')
				endwith
			endwith

* Set up the Delete section.

			loSection = .AddSection()
			with loSection
				.Caption = 'Delete'
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Ignore'
					.Image   = lcImagePath + 'ignore.png'
				endwith
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Clean Up'
					.Image   = lcImagePath + 'cleanup.png'
					.AddBar('\<Clean Up Conversation', , lcImagePath + 'cleanup.png')
					.AddBar('\<Clean Up Folder')
					.AddBar('\<Clean Up Folder & Subfolders')
				endwith
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Junk'
					.Image   = lcImagePath + 'junk.png'
					.AddBar('\<Block Sender', , lcImagePath + 'junk.png')
					.AddBar('Never Block \<Sender')
					.AddBar("Never Block Sender's \<Domain (@example.com)")
					.AddBar('Never Block this Group or \<Mailing List')
					.AddBar()
					.AddBar('\<Not Junk', , lcImagePath + 'newemailsmall.png', '.F.')
						&& .F. is passed as a string because it's an expression
						&& that's evaluated
					.AddBar('Junk E-mail \<Options...')
				endwith
				loButton = .AddButton()
				with loButton
					.Caption           = 'Delete'
					.Image             = lcImagePath + 'delete.png'
					.EnabledExpression = '.F.'
				endwith
				loButton = .AddButton()
				with loButton
					.Caption = 'Archive'
					.Image   = lcImagePath + 'archive.png'
				endwith
			endwith

* Set up the Move section.

			loSection = .AddSection()
			with loSection
				.Caption = 'Move'
				loButton = .AddButton(, 'SFRibbonToolbarButton')
				with loButton
					.Caption = 'Move'
					.Image   = lcImagePath + 'move.png'
					.AddBar('Inbox')
					.AddBar()
					.AddBar('\<Other Folder...')
					.AddBar('\<Copy to Folder...')
				endwith
				loButton = .AddButton(, 'SFRibbonToolbarButton')
				with loButton
					.Caption = 'Rules'
					.Image   = lcImagePath + 'rules.png'
					.AddBar('Always move messages from: whoever@whatever.com')
					.AddBar()
					.AddBar('Create R\<ule...')
					.AddBar('Manage Ru\<les & Alerts...')
				endwith

* Add a textbox to the section to show that the ribbon can support other types
* of controls besides buttons. We have to manually call CalculateWidth to
* adjust the section width.

				loControl = .AddControl('Test', 'textbox', '')
				with loControl
					.FontName = 'Segoe UI'
					.Width    = 200
					.Height   = 24
					.Top      = int((loSection.Height - .Height)/2)
						&& center it vertically
				endwith
				.CalculateWidth()
			endwith
		endwith

* Set up the Send/Receive tab.

		loTab = .AddTab()
		with loTab
			.Caption = 'Send / Receive'

* Set up the Send & Receive section.

			loSection = .AddSection()
			with loSection
				.Caption = 'Send & Receive'
				loButton = .AddButton()
				with loButton
					.Caption = 'Send/Receive' + chr(13) + 'All Folders'
					.Image   = lcImagePath + 'sendreceive.png'
				endwith
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Update Folder'
					.Image   = lcImagePath + 'updatefoldersmall.png'
				endwith
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Send All'
					.Image   = lcImagePath + 'sendallsmall.png'
				endwith
				loButton = .AddHorizontalButton()
				with loButton
					.Caption = 'Send/Receive Groups'
					.Image   = lcImagePath + 'updatefoldersmall.png'
				endwith
			endwith

* Set up the Download section.

			loSection = .AddSection()
			with loSection
				.Caption = 'Download'
				loButton = .AddButton()
				with loButton
					.Caption = 'Show' + chr(13) + 'Progress'
					.Image   = lcImagePath + 'showprogress.png'
				endwith
				loButton = .AddButton()
				with loButton
					.Caption = 'Cancel' + chr(13) + 'All'
					.Image   = lcImagePath + 'cancelall.png'
				endwith
			endwith
		endwith

* Add a shortcut menu to the ribbon.

		.AddBar('Change Theme to Colorful', ;
			"_screen.oRibbon.Theme = 'Colorful'" + chr(13) + "_screen.Refresh()", , ;
			"_screen.oRibbon.Theme <> 'Colorful'")
		.AddBar('Change Theme to Dark Grey', ;
			"_screen.oRibbon.Theme = 'Dark Grey'" + chr(13) + "_screen.Refresh()", , ;
			"_screen.oRibbon.Theme <> 'Dark Grey'")

* Auto-select the Home tab.

		.Home.Selected = .T.
	endwith
endif type('_screen.oRibbon') <> 'O'
