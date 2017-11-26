#!/usr/bin/perl
# Perl - v: 5.16.3
#------------------------------------------------------------------------------#
# ExtractFaceLang.pl  : Strings for ExtractFace
# WebSite             : http://le-tools.com/ExtractFace.html
# Documentation       : http://le-tools.com/ExtractFaceDoc.html
# SourceForge         : https://sourceforge.net/p/extractface
# GitHub              : https://github.com/arioux/ExtractFace
# Creation            : 2015-08-01
# Modified            : 2017-11-26
# Author              : Alain Rioux (admin@le-tools.com)
#
# Copyright (C) 2015-2017  Alain Rioux (le-tools.com)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#------------------------------------------------------------------------------#
# Modules
#------------------------------------------------------------------------------#
use strict;
use warnings;

#------------------------------------------------------------------------------#
sub loadStr
#------------------------------------------------------------------------------#
{
  # Local variables
  my ($refSTR, $LANG_FILE) = @_;
  # Open and load string values
  open(LANG, "<:encoding(UTF-8)", $LANG_FILE);
  my @tab = <LANG>;
  close(LANG);
  # Store values  
  foreach (@tab) {
    chomp($_);
    s/[^\w\=\s\.\!\,\-\)\(\']//g;
    my ($key, $value) = split(/ = /, $_);
    $value = encode("iso-8859-1", $value); # Revaluate with different language encoding
    if ($key) { $$refSTR{$key}  = $value; }
  }
  
}  #--- End loadStr

#------------------------------------------------------------------------------#
sub loadDefaultStr
#------------------------------------------------------------------------------#
{
  # Local variables
  my $refSTR = shift;
  
  # Set default strings
  
  # General strings
  $$refSTR{'Filename'}    = 'Filename';
  $$refSTR{'Directory'}   = 'Directory';
  $$refSTR{'SelectDir'}   = 'Select a directory';
  $$refSTR{'Select'}      = 'Select';
  $$refSTR{'Open'}        = 'Open';
  $$refSTR{'Dump'}        = 'Dump';
  $$refSTR{'DumpNow'}     = 'Dump now';
  $$refSTR{'AddToQueue'}  = 'Add to queue';
  $$refSTR{'Ok'}          = 'Ok';
  $$refSTR{'Include'}     = 'Include';
  $$refSTR{'Options'}     = 'Options';
  $$refSTR{'Cancel'}      = 'Cancel';
  $$refSTR{'cancelled'}   = 'cancelled';
  $$refSTR{'Opening'}     = 'Opening';
  $$refSTR{'Process'}     = 'Process';
  $$refSTR{'Processing'}  = 'Processing';
  $$refSTR{'Writing'}     = 'Writing';
  $$refSTR{'Enable'}      = 'Enable';
  $$refSTR{'Scroll'}      = 'Scroll';
  $$refSTR{'Scrolling'}   = 'Scrolling';
  $$refSTR{'Expand'}      = 'Expand';
  $$refSTR{'Remove'}      = 'Remove';
  $$refSTR{'Top'}         = 'Top';
  $$refSTR{'leftCol'}     = 'Left column';
  $$refSTR{'rightCol'}    = 'Right column';
  $$refSTR{'Bottom'}      = 'Bottom';
  $$refSTR{'Expanding'}   = 'Expanding';
  $$refSTR{'ScrollExpand'} = 'Scroll and Expand';
  $$refSTR{'Downloading'} = 'Downloading';
  $$refSTR{'thePage'}     = 'the page';
  $$refSTR{'theReport'}   = 'the report';
  $$refSTR{'Saving'}      = 'Saving';
  $$refSTR{'Parsing'}     = 'Parsing';
  $$refSTR{'Finishing'}   = 'Finishing';
  $$refSTR{'finished'}    = 'finished';
  $$refSTR{'Creating'}    = 'Creating';
  $$refSTR{'outputFile'}  = 'output file';
  $$refSTR{'sheet'}       = 'sheet';
  $$refSTR{'page'}        = 'page';
  $$refSTR{'Pages'}       = 'Pages';
  $$refSTR{'textFile'}    = 'text file';
  $$refSTR{'inProgress'}  = 'in progress';
  $$refSTR{'List'}        = 'List';
  $$refSTR{'Category'}    = 'Category';
  $$refSTR{'Lists'}       = 'Lists';
  $$refSTR{'ProfileID'}   = 'Profile ID';
  $$refSTR{'Image'}       = 'Image';
  $$refSTR{'url'}         = 'URL';
  $$refSTR{'Name'}        = 'Name';
  $$refSTR{'Details'}     = 'Details';
  $$refSTR{'Path'}        = 'Path';
  $$refSTR{'imgPath'}     = 'Image Path';
  $$refSTR{'imgPath2'}    = 'Image Path or URL';
  $$refSTR{'originURL'}   = 'Origin URL';
  $$refSTR{'eventURL'}    = 'Event URL';
  $$refSTR{'Count'}       = 'Count';
  $$refSTR{'Date'}        = 'Date';
  $$refSTR{'Dates'}       = 'Dates';
  $$refSTR{'Start'}       = 'Start';
  $$refSTR{'Time'}        = 'Time';
  $$refSTR{'End'}         = 'End';
  $$refSTR{'Message'}     = 'Message';
  $$refSTR{'files'}       = 'files';
  $$refSTR{'profileIcons'}    = 'profile icons';
  $$refSTR{'Help'}            = 'Help';
  $$refSTR{'Quit'}            = 'Quit';
  $$refSTR{'Warning'}         = 'Warning';
  $$refSTR{'warn2'}           = 'You must select at least one option.';
  $$refSTR{'warn3'}           = 'You are not in the right page.';
  $$refSTR{'warn4'}           = 'You are not on Facebook.';
  $$refSTR{'Error'}           = 'Error';
  $$refSTR{'errMozRepl'}      = 'You must start Firefox MozRepl add-on.';
  $$refSTR{'processRunning'}  = 'A process is already running. Wait until it stops or restart the program.';
  $$refSTR{'processCrash'}    = 'Process crash';
  $$refSTR{'crash'}           = 'Crashed, ExtractFace will try to resume';
  $$refSTR{'browseFolder'}    = 'Browse folder in Explorer';
  $$refSTR{'resumeProcess'}   = 'Resuming process';
  $$refSTR{'Progress'}        = 'Progress';
  $$refSTR{'ReloadPage'}      = 'Reload the page';
  $$refSTR{'noProfileDumped'} = 'No profile were dumped';
  $$refSTR{'AutoScroll'}      = "Auto scroll";
  $$refSTR{'NotFound'}        = 'Not found';
  # Dump Windows
  $$refSTR{'currProfileID'} = 'Current Profile ID';
  $$refSTR{'errProfileID'}  = 'Profile ID not found.';
  $$refSTR{'Albums'}        = 'Albums';
  $$refSTR{'albumNames'}    = 'Album name';
  $$refSTR{'albumURLs'}     = 'Album url';
  $$refSTR{'albumID'}       = 'Album ID';
  $$refSTR{'loadAlbum'}     = 'Loading the album page';
  $$refSTR{'loadAlbumFail'} = 'No album title found.';
  $$refSTR{'chPublishDate'} = 'Publication date';
  $$refSTR{'openAlbumDir'}  = 'Open album folder';
  $$refSTR{'SmallPic'}      = 'Small pictures';
  $$refSTR{'LargePic'}      = 'Large pictures';
  $$refSTR{'browsePicPage'} = 'Browsing picture or video page';
  $$refSTR{'picsAndVids'}   = 'pictures and/or videos';
  $$refSTR{'pictureID'}     = 'Picture/Video ID';
  $$refSTR{'picturePage'}   = 'Picture/Video Page URL';
  $$refSTR{'smallPicURL'}   = 'Small picture URL';
  $$refSTR{'largePicURL'}   = 'Large picture URL';
  $$refSTR{'videoURL'}      = 'Video URL';
  $$refSTR{'dumpAlbumError'} = 'Some pictures/videos have not been downloaded.';
  $$refSTR{'friends'}        = 'Friends';
  $$refSTR{'gatherFriendsLists'} = 'Gathering friend lists';
  $$refSTR{'MutualFriends'} = 'Mutual friends';
  $$refSTR{'EventMembers'}  = 'Event members';
  $$refSTR{'gatherEvent'}   = 'Gathering event details';
  $$refSTR{'EventDataPage'} = 'Event data page';
  $$refSTR{'contributors'}  = 'Contributors';
  $$refSTR{'Types'}         = 'Types';
  $$refSTR{'Comments'}      = 'Comments';
  $$refSTR{'Likes'}         = 'Likes';
  $$refSTR{'VPosts'}        = 'Visitor Posts';
  $$refSTR{'EventPosts'}    = 'Event Posts';
  $$refSTR{'LikesPage'}     = 'Likes page';
  $$refSTR{'groupMembers'}  = 'Group Members';
  $$refSTR{'loadOlderMsg'}  = 'Load Older Messages';
  $$refSTR{'loadNewerMsg'}  = 'Load Newer Messages';
  $$refSTR{'theConv'}       = 'the conversation';
  $$refSTR{'Chat'}          = 'Chat';
  $$refSTR{'Messenger'}     = 'Messenger';
  $$refSTR{'Me'}            = 'Me';
  $$refSTR{'AttachedDoc'}   = 'Attached document';
  $$refSTR{'Images'}        = 'Images';
  $$refSTR{'Pictures'}      = 'Pictures';
  $$refSTR{'Videos'}        = 'Videos';
  $$refSTR{'vocalMsg'}      = 'Vocal messages';
  $$refSTR{'All'}           = 'All';
  $$refSTR{'Range'}         = 'Range';
  $$refSTR{'CurrChat'}      = 'Current chat';
  $$refSTR{'SelectChat'}    = 'Select chat(s)';
  $$refSTR{'SharedLink'}    = 'Shared link';
  $$refSTR{'Video'}         = 'Video';
  $$refSTR{'noMsgDumped'}   = 'No message were dumped';
  $$refSTR{'noVocalMsgDumped'} = 'No vocal message were dumped';
  $$refSTR{'MobileFacebook'} = 'Mobile Facebook';
  $$refSTR{'Searching'}     = 'Searching';
  $$refSTR{'browseAllChat'} = 'Browsing all messages';
  $$refSTR{'vocalMsgLast'}  = 'Vocal message, last';
  $$refSTR{'Listen'}        = 'Listen';
  $$refSTR{'Contacts'}      = 'Contacts';
  $$refSTR{'CurrPage'}      = 'Current page';
  $$refSTR{'PicPages'}      = 'Picture pages (People profile only)';
  $$refSTR{'browseAllPagesURLs'} = 'Browsing all pages';
  # Queue
  $$refSTR{'ShowQueue'}       = 'Show queue';
  $$refSTR{'Queue'}           = 'Queue';
  $$refSTR{'MoveUp'}          = 'Move up';
  $$refSTR{'MoveDown'}        = 'Move down';
  $$refSTR{'GoToPage'}        = 'Go to page';
  $$refSTR{'Delete'}          = 'Delete';
  $$refSTR{'addingQueue'}     = 'Creating and adding processes to queue';
  $$refSTR{'addedQueue'}      = 'Process(es) added to queue';
  $$refSTR{'errAddQueue'}     = 'Error(s) while adding process(es) to queue';
  $$refSTR{'pendingJob'}      = 'There are pending jobs in queue, load them';
  $$refSTR{'pendingJobWarn'}  = 'Note: If you select No, all pending jobs will be deleted';
  $$refSTR{'queueExists'}     = 'There is job with the same filename in queue. If you continue, the report may be replaced. Continue anyway';
  # Config
  $$refSTR{'Settings'}        = 'Settings';
  $$refSTR{'lblGenOpt'}       = 'General';
  $$refSTR{'Tool'}            = 'Tool';
  $$refSTR{'btnExportLang'}   = 'Export Lang.ini';
  $$refSTR{'chAutoUpdate'}    = 'Check for update at startup';
  $$refSTR{'Functions'}       = 'Functions';
  $$refSTR{'DynamicMenu'}     = 'Dynamic menu';
  $$refSTR{'chRememberPos'}   = 'Remember position of all windows';
  $$refSTR{'lblTimeToWait'}   = 'Time for loading';
  $$refSTR{'seconds'}         = 'seconds';
  $$refSTR{'Charset'}         = 'Charset';
  $$refSTR{'Logging'}         = 'Logging';
  $$refSTR{'OpenLog'}         = 'Open log file';
  $$refSTR{'ClearLog'}        = 'Clear log file';
  $$refSTR{'tfTimeToWaitTip'} = 'When loading page or scrolling, time to wait before any action. Increase this time for more stability. Default is 2.';
  $$refSTR{'lblMaxLoading'}   = 'Max scrolling (chat)';
  $$refSTR{'lblMaxScroll'}    = 'Max scrolling (other)';
  $$refSTR{'ByPage'}                = 'By page';
  $$refSTR{'tfMaxScrollByPageTip'}  = 'Stop scrolling after a maximum of pages displayed. Default is 0 (No maximum).';
  $$refSTR{'ByDate'}                = 'By date';
  $$refSTR{'tfMaxScrollByDateTip'}  = 'Stop scrolling when the given date is reached.';
  $$refSTR{'chOptSeemore'}    = 'See more (include comments and replies)';
  $$refSTR{'chOptPosts'}      = 'More posts';
  $$refSTR{'chOptTranslate'}  = 'See translation';
  $$refSTR{'WhenLoading'}     = 'When loading';
  $$refSTR{'AutoLoadScroll'}  = 'Load and scroll automatically';
  $$refSTR{'chOptScrollTop'}  = 'Scroll back to top when loaded';
  $$refSTR{'RememberSaveDir'} = 'Remember folder used for report';
  $$refSTR{'WhenProcessing'}  = 'When processing';
  $$refSTR{'SilentProgress'}  = 'Silent progression when using queue';
  $$refSTR{'WhenFinished'}    = 'When finished';
  $$refSTR{'OpenReport'}      = 'Open report';
  $$refSTR{'DontOpenReport'}  = 'Disable opening when using queue';
  $$refSTR{'chCloseUsedTabs'} = 'Close used tabs';
  $$refSTR{'chDelTempFiles'}  = 'Delete temp files';
  # Update Window
  $$refSTR{'update1'}       = 'You have the latest version installed.';
  $$refSTR{'update2'}       = 'Check for update';
  $$refSTR{'update3'}       = 'Update';
  $$refSTR{'update4'}       = 'Version';
  $$refSTR{'update5'}       = 'is available. Download it';
  $$refSTR{'errConnection'} = 'Error connection';
  $$refSTR{'returnedCode'}  = 'Returned code';
  $$refSTR{'returnedError'} = 'Returned error';
  # About
  $$refSTR{'about'}             = 'About';
  $$refSTR{'author'}            = 'Author';
  $$refSTR{'translatedBy'}      = 'Translated by';
  $$refSTR{'website'}           = 'Website';
  $$refSTR{'translatorName'}    = '-';
  $$refSTR{'chStartMinimized'}  = "Don't show this window on startup";
  $$refSTR{'lblText4'}          = 'Use taskbar icon to access functions';

}  #--- End loadStrings

#------------------------------------------------------------------------------#
1;