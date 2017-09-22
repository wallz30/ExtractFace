#!/usr/bin/perl
# Perl - v: 5.16.3
#------------------------------------------------------------------------------#
# ExtractFaceFunctions.pl : Functions for ExtractFace
# WebSite                 : http://le-tools.com/ExtractFace.html
# Documentation           : http://le-tools.com/ExtractFaceDoc.html
# SourceForge             : https://sourceforge.net/p/extractface
# GitHub                  : https://github.com/arioux/ExtractFace
# Creation                : 2015-08-01
# Modified                : 2017-09-22
# Author                  : Alain Rioux (admin@le-tools.com)
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

#------------------------------------------------------------------------------#
# Modules
#------------------------------------------------------------------------------#
use strict;
use warnings;
use Win32::GUI;
use Win32::GUI::Grid;
use Win32::Process;
use Time::HiRes qw/gettimeofday/;

#------------------------------------------------------------------------------#
# Global variables
#------------------------------------------------------------------------------#
my $URL_VER      = 'http://www.le-tools.com/download/ExtractFaceVer.txt';      # Url of the version file
my $URL_TOOL     = 'http://le-tools.com/ExtractFace.html#Download';            # Url of the tool (download)

#--------------------------#
sub winPIDThr
#--------------------------#
{
  my ($refTHR, $refARROW, $refHOURGLASS, $refWinPID, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  # Deal with crash
  $SIG{__DIE__} = sub {
    my $msgErr = $_[0];
    chomp($msgErr);
    $msgErr =~ s/[\t\r\n]/ /g;
    if ($msgErr =~ /NS_ERROR_FILE_IS_LOCKED/) { # Restart a new thread to continue
      $$refTHR = threads->create(\&winPIDThr, $refTHR, $refARROW, $refHOURGLASS, $refWinPID,
                                 $DEBUG_FILE, $refCONFIG, $refWin, $refSTR);
    } else {
      &debug($msgErr, $DEBUG_FILE) if $$refCONFIG{'DEBUG_LOGGING'};
			my $err = (split(/ at /, $msgErr))[0];
      $$refWin->ChangeCursor($$refARROW);
      $$refWin->Tray->Change(-tip => 'ExtractFace');
      Win32::GUI::MessageBox($$refWinPID, "$$refSTR{'processCrash'}: $err", $$refSTR{'Error'}, 0x40010);
			threads->exit();
    }
  };
  my $mech;
  eval { $mech = WWW::Mechanize::Firefox->new(tab => 'current'); };
  if ($@) {
    Win32::GUI::MessageBox($$refWinPID, $$refSTR{'errMozRepl'}, $$refSTR{'Error'}, 0x40010)
    if $@ =~ /Failed to connect to/;
    threads->exit();
  }
  $$refWin->ChangeCursor($$refHOURGLASS);
	# Get the profile ID
  my ($pidCode, $pageType) = &getCurrPidCode(\$mech);
	if ($pidCode) { $$refWinPID->tfPIDTitle->Text($pidCode); }
  else          { $$refWinPID->tfPIDTitle->Text($$refSTR{'errProfileID'}); }
  $$refWin->ChangeCursor($$refARROW);

}  #--- End winPIDThr

#--------------------------#
sub getCurrPidCode
#--------------------------#
{
  # Local variables
  my $refMech = shift;
  my $pidCode = undef;
  # Second return value (pageType): 0 = unknown, 1 = People, 2 = Groups, 3 = Pages (Business), 4 = In Messenger
	# Normal profile
	if ($pidCode = $$refMech->selector('div._4a8n a', any => 1)) {
		if ($pidCode->{outerHTML} =~ /user.php\?id=(\d+)/) { return($1, 1); }
	} elsif ($pidCode = $$refMech->selector('a.profilePicThumb', any => 1)) {
		if ($pidCode->{href} =~ /profile_id=(\d+)/ or $pidCode->{href} =~ /fbid=(\d+)/) { return($1, 1); }
	# Group profile
	} elsif ($pidCode = $$refMech->selector('a.coverImage', any => 1)) {
		if ($pidCode->{outerHTML} =~ /referrer_profile_id=(\d+)/) { return($1, 2); }
	# Business page
	} elsif ($pidCode = $$refMech->selector('a._2dgj', any => 1)) {
		if ($pidCode->{href} =~ /\/([^\/]+)\/photos/) { return($1, 3); }
  # In Messenger
  } elsif ($$refMech->uri =~ /\/messages/) {
    my $currURL = $$refMech->uri;
    if ($currURL =~ /facebook.com\/messages\/t\/search\/([^\/\?]+)\/?/ or
        $currURL =~ /facebook.com\/messages\/t\/([^\/\?]+)\/?/         or
        $currURL =~ /facebook.com\/messages\/archived\/t\/([^\/\?]+)\/?/) {
      $pidCode = $1;
    }
    return($pidCode, 4);
  # Event page
  } elsif ($$refMech->uri =~ /\/events((?:\/\d+))?/) {
    if ($1 and $1 =~ /\/(\d+)/) { $pidCode = $1; }
    return($pidCode, 5);
  # Mutual Friends
  } elsif ($$refMech->uri =~ /\/mutual_friends/) {
    if ($$refMech->uri =~ /uid=(\d+)&node=(\d+)/) {
      $pidCode = "$1 - $2";
    }
    return($pidCode, 6);
  }
  return(0, 0);
  
}  #--- End getCurrPidCode

#--------------------------#
sub loadDumpPageThr
#--------------------------#
{
  # Local variables
  my ($refTHR, $refARROW, $refHOURGLASS, $refWinDump, $typeDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  # $typeDump: 1=Album, 2=Friends, 3=MutualFriends, 4=Event, 5=Contrib, 6=GroupMembers, 7=Chat, 8=Contacts
  # Deal with crash
  $SIG{__DIE__} = sub {
    my $msgErr = $_[0];
    chomp($msgErr);
    $msgErr =~ s/[\t\r\n]/ /g;
    if ($msgErr =~ /NS_ERROR_FILE_IS_LOCKED/) { # Restart a new thread to continue
      $$refTHR = threads->create(\&loadDumpPageThr, $refTHR, $refARROW, $refHOURGLASS, $refWinDump, $typeDump,
                                 $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR);
    } else {
      &debug($msgErr, $DEBUG_FILE) if $$refCONFIG{'DEBUG_LOGGING'};
			my $err = (split(/ at /, $msgErr))[0];
      $$refWin->ChangeCursor($$refARROW);
      $$refWin->Tray->Change(-tip => 'ExtractFace');
      Win32::GUI::MessageBox($$refWinDump, "$$refSTR{'processCrash'}: $err", $$refSTR{'Error'}, 0x40010);
			threads->exit();
    }
  };
  my $mech;
  eval { $mech = WWW::Mechanize::Firefox->new(tab => 'current'); };
  if ($@) {
    Win32::GUI::MessageBox($$refWinDump, $$refSTR{'errMozRepl'}, $$refSTR{'Error'}, 0x40010) if $@ =~ /Failed to connect to/;
    threads->exit();
  }
  # Valid current page
  if ($mech->uri() =~ /facebook.com/) {
    $$refWinDump->ChangeCursor($$refHOURGLASS);
    if    ($typeDump == 1) { &loadDumpAlbum(        \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 2) { &loadDumpFriends(      \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 3) { &loadDumpMutualFriends(\$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 4) { &loadDumpEventMembers( \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 5) { &loadDumpContrib(      \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 6) { &loadDumpGroupMembers( \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 7) { &loadDumpChat(         \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    elsif ($typeDump == 8) { &loadDumpContacts(     \$mech, $refWinDump, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR); }
    $$refWinDump->lblInProgress->Text('');
    $$refWinDump->ChangeCursor($$refARROW);
  } else { Win32::GUI::MessageBox($$refWinDump, $$refSTR{'warn4'}, $$refSTR{'Error'}, 0x40010); }
  if    ($typeDump == 1) { &isDumpAlbumsReady(        $refWinDump); }
  elsif ($typeDump == 2) { &isDumpFriendsReady(       $refWinDump); }
  elsif ($typeDump == 3) { &isDumpMutualFriendsReady( $refWinDump); }
  elsif ($typeDump == 4) { &isDumpEventMembersReady(  $refWinDump); }
  elsif ($typeDump == 5) { &isDumpContribReady(       $refWinDump); }
  elsif ($typeDump == 6) { &isDumpGroupMembersReady(  $refWinDump); }
  elsif ($typeDump == 7) { &isDumpChatReady(          $refWinDump); }
  elsif ($typeDump == 8) { &isDumpContactsReady(      $refWinDump); }
  
}  #--- End loadDumpPageThr
  
#--------------------------#
sub loadDumpAlbum
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinAlbums, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  # Valid current page
  my ($validPage, $currURL) = &validAlbumPage($refMech, $refWinAlbums, $refCONFIG, $refSTR);
  if ($validPage) {
    # Get album names and urls
    my %albums;
    my $pageType = $$refWinAlbums->tfPageType->Text();
    if    ($pageType == 1 or $pageType == 2) { # People or Group
      &getListAlbums($refMech, $refWinAlbums, $pageType, $USERDIR, $refCONFIG);
      if ($pageType == 2) {
        my $nextAlbumPage = $$refMech->selector('a.next.uiButton.uiButtonNoText', any => 1);
        while ($nextAlbumPage) {
          if ($nextAlbumPage->{outerHTML} =~ /uiButtonDisabled/) { last; }
          else {
            $nextAlbumPage->click();
            sleep($$refCONFIG{'TIME_TO_WAIT'});
            &getListAlbums($refMech, $refWinAlbums, $pageType, $USERDIR, $refCONFIG);
            $nextAlbumPage = $$refMech->selector('a.next.uiButton.uiButtonNoText', any => 1);
          }
        }
        # Videos ?
        my @headerLinks = $$refMech->selector('li._dur a._duq');
        foreach (@headerLinks) {
          if (my $code = $_->{outerHTML}) {
            if ($code =~ /\/videos/ and $code =~ /href=\"([^\"]+)\"/) {
              my $url = $1;
              if ($code =~ /<a[^\>]+\>([^\<]+)\</) {
                my $name = $1;
                $url     = 'https://www.facebook.com'.$url if $url !~ /^http/;
                if (my $i = $$refWinAlbums->GridAlbums->InsertRow($name, -1)) {
                  $$refWinAlbums->GridAlbums->SetCellText($i, 0, ''        );
                  $$refWinAlbums->GridAlbums->SetCellType($i, 0, GVIT_CHECK);
                  $$refWinAlbums->GridAlbums->SetCellCheck($i, 0, 1);
                  $$refWinAlbums->GridAlbums->SetCellText($i, 1, $name   );
                  $$refWinAlbums->GridAlbums->SetCellText($i, 2, 'Videos');
                  $$refWinAlbums->GridAlbums->SetCellText($i, 3, $url    );
                  $$refWinAlbums->GridAlbums->Refresh();
                }
              }
            }
          }
        }
      }
    } elsif ($pageType == 3) { # Page (Business)
      &getListAlbums($refMech, $refWinAlbums, $pageType, $USERDIR, $refCONFIG);
    }
    $$refWinAlbums->tfAlbumCurrURL->Text($currURL);
    if ($$refWinAlbums->GridAlbums->GetRows() > 1) {
      # Feed the Album Grid
      $$refWinAlbums->GridAlbums->SetCellCheck(0, 0, 1);
      $$refWinAlbums->GridAlbums->Refresh();
      $$refWinAlbums->GridAlbums->AutoSize();
      $$refWinAlbums->GridAlbums->ExpandLastColumn();
      $$refWinAlbums->GridAlbums->BringWindowToTop();
    } else { Win32::GUI::MessageBox($$refWinAlbums, $$refSTR{'loadAlbumFail'}, $$refSTR{'Error'}, 0x40010); }
  } else { Win32::GUI::MessageBox($$refWinAlbums, $$refSTR{'warn3'}, $$refSTR{'Error'}, 0x40010); }

}  #--- End loadDumpAlbum

#--------------------------#
sub loadDumpFriends
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinFriends, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  # Valid current page
  my $saveDir = $$refWinFriends->tfDirSaveFriends->Text();
  $$refWinFriends->lblInProgress->Text($$refSTR{'gatherFriendsLists'}.'...');
  my $currURL = $$refMech->uri();
  chop($currURL) if $currURL =~ /#$/;
  my $currTitle;
  if ($currURL !~ /\/friends\/?/ and $currURL !~ /sk=friends/) {
    if ($$refCONFIG{'AUTO_LOAD_SCROLL'}) {
      # Trying to get the good page
      if ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/) {
        my $profilID = $1;
        my $goodURL = "https://www.facebook.com/profile.php?id=$profilID&sk=friends";
        ($currURL, $currTitle) = &loadPage($refMech, $goodURL, $$refCONFIG{'TIME_TO_WAIT'});
      } elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/) {
        my $goodURL = "https://www.facebook.com/$1/friends";
        ($currURL, $currTitle) = &loadPage($refMech, $goodURL, $$refCONFIG{'TIME_TO_WAIT'});
      }
      # Re evaluate current page
      if (($currURL !~ /\/friends\/?$/ and $currURL !~ /sk=friends/) or $currTitle =~ /Page Not Found/) {
        Win32::GUI::MessageBox($$refWinFriends, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010);
      } else {
        if    ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/) { $currTitle = $1; }
        elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/                ) { $currTitle = $1; }
        $currTitle =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g;
        $$refWinFriends->tfFriendName->Text("$currTitle - $$refSTR{'friends'}");
      }
    } else { Win32::GUI::MessageBox($$refWinFriends, $$refSTR{'warn3'}, $$refSTR{'Error'}, 0x40010); }
  # You are in the right page
  } else {
    if    ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/) { $currTitle = $1; }
    elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/                ) { $currTitle = $1; }
    $currTitle =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g;
    $$refWinFriends->tfFriendName->Text("$currTitle - $$refSTR{'friends'}");
  }
  # List available categories
  if ($$refWinFriends->tfFriendName->Text()) {
    $$refWinFriends->tfFriendCurrURL->Text($currURL);
    my %categories;
    my $header = $$refMech->selector('div._3dc.lfloat._ohe._5brz', one => 1)->{innerHTML};
    my @div    = split(/\>/, $header);
    foreach (@div) {
      if (/href=\"([^\"]+)\"/) {
        my $url  = $1;
        $url =~ s/&amp;/&/g;
        my $catName;
        if (/name=\"([^\"]+)\"/) { $catName = $1; }
        if ($catName) {
          my $cat = encode($$refCONFIG{'CHARSET'}, $catName);
          $cat =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g;
          $categories{$cat}{name} = $catName;
          $categories{$cat}{url}  = $url;
          if (/aria-controls=\"([^\"]+)\"/) {
            my $catId = $1;
            $categories{$cat}{catId} = $catId;
          }
        }
      }
    }
    foreach my $cat (sort keys %categories) {
      if ($categories{$cat} and my $i = $$refWinFriends->GridFriends->InsertRow($cat, -1)) {
        $$refWinFriends->GridFriends->SetCellText($i, 0, ''        );
        $$refWinFriends->GridFriends->SetCellType($i, 0, GVIT_CHECK);
        $$refWinFriends->GridFriends->SetCellCheck($i, 0, 1);
        $$refWinFriends->GridFriends->SetCellText($i, 1, $cat                    );
        $$refWinFriends->GridFriends->SetCellText($i, 2, $categories{$cat}{catId});
        $$refWinFriends->GridFriends->SetCellText($i, 3, $categories{$cat}{url}  );
      }
    }
    $$refWinFriends->GridFriends->AutoSize();
    $$refWinFriends->GridFriends->ExpandLastColumn();
    $$refWinFriends->GridFriends->Refresh();
    $$refWinFriends->GridFriends->BringWindowToTop();
  }

}  #--- End loadDumpFriends

#--------------------------#
sub loadDumpMutualFriends
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinMutualFriends, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $currURL = $$refMech->uri();
  if ($currURL =~ /uid=(\d+)&node=(\d+)/) {
    my $id1 = $1;
    my $id2 = $2;
    $$refWinMutualFriends->tfMutualFriendsName->Text("$id1 - $id2 - $$refSTR{'MutualFriends'}");
    $$refWinMutualFriends->tfMutualFriendsCurrURL->Text($$refMech->uri);
  } else {
    $$refWinMutualFriends->btnMutualFriendsDumpNow->Disable();
    $$refWinMutualFriends->btnMutualFriendsAddQueue->Disable();
    Win32::GUI::MessageBox($$refWinMutualFriends, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010);
  }

}  #--- End loadDumpMutualFriends

#--------------------------#
sub loadDumpContrib
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinContrib, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $title;
  my $currURL = $$refMech->uri();
  my ($pidCode, $pageType) = &getCurrPidCode($refMech);
  if    ($pidCode) { $title = $pidCode; }
  elsif ($currURL =~ /profile.php\?id=([^\/\&\#]+)/                  ) { $title = $1; }
  elsif ($currURL =~ /fbid=([^\/\&\#]+)/                             ) { $title = $1; }
  elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/) { $title = $1; }
  $$refWinContrib->tfContribName->Text("$title - $$refSTR{'contributors'}");
  $$refWinContrib->tfStartURL->Text($currURL);
  if (my ($pidCode, $pageType) = &getCurrPidCode($refMech)) { $$refWinContrib->tfContribID->Text($pidCode); }
  if ($currURL =~ /posts_to_page/) { # Visitor posts popup is open
    $$refWinContrib->chContribVPosts->Enable();
    $$refWinContrib->chContribEventPosts->Disable();
  } elsif ($pageType == 5) { # On Event Page
    $$refWinContrib->chContribEventPosts->Enable();
    $$refWinContrib->chContribVPosts->Disable();
  } else { # Other pages
    $$refWinContrib->chContribVPosts->Disable();
    $$refWinContrib->chContribEventPosts->Disable();
  }

}  #--- End loadDumpContrib

#--------------------------#
sub loadDumpEventMembers
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinEvent, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $saveDir = $$refWinEvent->tfDirSaveEvent->Text();
  my $tempDir = "$USERDIR\\temp";
  # Valid current page
  my $currURL = $$refMech->uri();
  my $idEvent;
  if ($currURL =~ /https:\/\/(?:www|web).facebook.com\/events\/(\d+)\/?/) {
    $idEvent = $1;
    $$refWinEvent->tfEventCurrURL->Text($currURL);
    $$refWinEvent->lblInProgress->Text($$refSTR{'gatherEvent'}.'...');
    # Gather event details
    my $eventDetailsURL = "https://www.facebook.com/events/ajax/guest_list/?acontext[ref]=51&acontext[source]=1&acontext[action_history]=[{%22surface%22%3A%22permalink%22%2C%22mechanism%22%3A%22surface%22%2C%22extra_data%22%3A[]}]&event_id=$idEvent&initial_tab=going&__pc=EXP1%3ADEFAULT&__asyncDialog=6&__a=1";
    mkdir("$tempDir") if !-d "$tempDir";
    my $localDataFile = "$tempDir\\data.txt";
    my $mechData = WWW::Mechanize::Firefox->new(create => 1, autodie => 0, );
    $mechData->get($eventDetailsURL, synchronize => 0);
    sleep($$refCONFIG{'TIME_TO_WAIT'});
    my $status = $mechData->save_content($localDataFile, $tempDir);
    while ($status->{currentState} != $status->{PERSIST_STATE_FINISHED}) { sleep(1); }
    # Parse html page
    open(my $tmp, $localDataFile);
    my $file_as_string = do { local $/ = <$tmp> };
    $file_as_string    =~ s/[\r\n]//g;
    close($tmp);
    # Event ID
    if ($file_as_string =~ /"eventID":"([^\"]+)"/ ) { $$refWinEvent->tfEventFilename->Text("$1 - $$refSTR{'EventMembers'}"); }
    # URL of content
    if ($file_as_string =~ /"prefetchURI":"([^\"]+)"/) {
      my $dataURL = $1;
      $dataURL =~ s#\\\/#\/#g;
      $dataURL =~ s/&amp;/&/g;
      $dataURL =~ s/\\u([[:xdigit:]]{4})/chr(hex $1)/eg;
      $dataURL = "https://www.facebook.com" . $dataURL . "&__pc=EXP1%3ADEFAULT&__a=1&__req=a";
      $$refWinEvent->tfDataURL->Text($dataURL);
    }
    # Guest list name
    my %glText;
    if ($file_as_string =~ /"typeaheadSubtitles":{([^\}]+)}/) {
      my $data = $1;
      my @lists = split(/,/, $data);
      foreach (@lists) {
        my ($lname, $ltext) = split(/:/, $_);
        $lname =~ s/"//g;
        $ltext =~ s/"//g;
        $glText{$lname} = $ltext;
      }
    }
    # Guest list members count
    my %glCount;
    if ($file_as_string =~ /"adminRSVPNuxCount":0,"counts":{([^\}]+)}/) {
      my $data = $1;
      my @lists = split(/,/, $data);
      foreach (@lists) {
        my ($lname, $lcount) = split(/:/, $_);
        $lname =~ s/"//g;
        $glCount{$lname} = $lcount;
      }
    }
    close($tmp);
    # Adjust Guest list checkbox text
    if ($glCount{going}    and $glText{going}   ) {
      my $encodedStr = encode($$refCONFIG{'CHARSET'}, "$glText{going} [$glCount{going}]");
      $encodedStr =~ s/\\u([[:xdigit:]]{4})/chr(hex $1)/eg;
      $$refWinEvent->chGoing->Text($encodedStr);
    }
    if ($glCount{maybe}    and $glText{maybe}   ) {
      my $encodedStr = encode($$refCONFIG{'CHARSET'}, "$glText{maybe} [$glCount{maybe}]");
      $encodedStr =~ s/\\u([[:xdigit:]]{4})/chr(hex $1)/eg;
      $$refWinEvent->chMaybe->Text($encodedStr);
    }
    if ($glCount{invited}  and $glText{invited} ) {
      my $encodedStr = encode($$refCONFIG{'CHARSET'}, "$glText{invited} [$glCount{invited}]");
      $encodedStr =~ s/\\u([[:xdigit:]]{4})/chr(hex $1)/eg;
      $$refWinEvent->chInvited->Text($encodedStr);
    }
    if ($glCount{declined} and $glText{declined}) {
      my $encodedStr = encode($$refCONFIG{'CHARSET'}, "$glText{declined} [$glCount{declined}]");
      $encodedStr =~ s/\\u([[:xdigit:]]{4})/chr(hex $1)/eg;
      $$refWinEvent->chDeclined->Text($encodedStr);
    }
    # Delete temporary files
    remove_tree($tempDir) if $$refCONFIG{'DEL_TEMP_FILES'};
  } else { Win32::GUI::MessageBox($$refWinEvent, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010); }

}  #--- End loadDumpEventMembers

#--------------------------#
sub loadDumpGroupMembers
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinGroupMembers, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $saveDir = $$refWinGroupMembers->tfDirSaveGroupMembers->Text();
  my $currURL = $$refMech->uri();
  my $type    = 0;
  my $idGroup;
  my $newURL;
  # Valid current page
  if ($currURL =~ /https:\/\/(?:www|web).facebook.com\/groups\/(\w+)/) {
    $idGroup = $1;
    my $load = 1;
    if (my $pageSel = $$refMech->selector('div.uiHeader.uiHeaderTopAndBottomBorder.uiHeaderSection a', any => 1)) {
      if ($currURL = $pageSel->{href}) {
        $currURL = 'https://www.facebook.com'.$currURL if $currURL !~ /^http/;
        $type = 1;
      }
    } elsif (my @menu = $$refMech->selector('div._2yaa')) {
      foreach (@menu) {
        if ($_->{outerHTML} =~ /data-key="members"/) {
          chop($currURL) if $currURL =~ /\/$/;
          if ($currURL !~ /members$/ and $currURL !~ /admins$/) { $currURL = $currURL . '/members'; }
          else 																									{ $load = 0; $newURL  = $currURL; 	}
          $type = 2;
        }
      }
    }
    if ($type and $currURL and $load and $$refCONFIG{'AUTO_LOAD_SCROLL'}) {
      $$refMech->get($currURL, synchronize => 0);
      sleep($$refCONFIG{'TIME_TO_WAIT'});
      $newURL = sprintf($$refMech->uri());
    }
  } elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/browse\/group_members\/\?gid=(\w+)/) {
    $idGroup = $1;
    $type    = 1;
    $newURL  = $currURL;
  }
  # Right page, get the values
  if ($type and $newURL and $idGroup and ($type == 1 and   $newURL =~ /https:\/\/(?:www|web).facebook.com\/browse\/group_members\/\?gid=$idGroup/) or
                                         ($type == 2 and (($newURL =~ /https:\/\/(?:www|web).facebook.com\/groups\/$idGroup\/members\/?/) or
                                          ($newURL =~ /https:\/\/(?:www|web).facebook.com\/groups\/$idGroup\/admins\/?/)))) {
    $$refWinGroupMembers->tfGroupMembersType->Text($type);
    $$refWinGroupMembers->tfGroupMembersName->Text("$idGroup - $$refSTR{'groupMembers'}");
    $$refWinGroupMembers->tfGroupMembersCurrURL->Text($newURL);
    my %categories;
    # Type with single list
    if ($type == 1) { $categories{$$refSTR{'groupMembers'}}{url} = $newURL; }
    # Type with tab selection
    else {
      if (my $membersPageName = $$refMech->selector('a._5bv4', any => 1)) {
        if ($membersPageName->{innerHTML} =~ /^(\w+)/) { $categories{$1}{url} = $membersPageName->{href}; }
      }
      if (my $adminsPageName  = $$refMech->selector('a._5bv3', any => 1)) {
        if ($adminsPageName->{innerHTML}  =~ /^(\w+)/) { $categories{$1}{url} = $adminsPageName->{href};  }
      }
    }
    foreach my $cat (sort keys %categories) {
      if (my $i = $$refWinGroupMembers->GridGroupMembers->InsertRow($cat, -1)) {
        $$refWinGroupMembers->GridGroupMembers->SetCellText($i, 0, ''        );
        $$refWinGroupMembers->GridGroupMembers->SetCellType($i, 0, GVIT_CHECK);
        $$refWinGroupMembers->GridGroupMembers->SetCellCheck($i, 0, 1);
        $categories{$cat}{name} = encode($$refCONFIG{'CHARSET'}, $cat);
        $$refWinGroupMembers->GridGroupMembers->SetCellText($i, 1, $categories{$cat}{name});
        $$refWinGroupMembers->GridGroupMembers->SetCellText($i, 2, $categories{$cat}{url});
        $$refWinGroupMembers->GridGroupMembers->AutoSize();
        $$refWinGroupMembers->GridGroupMembers->ExpandLastColumn();
        $$refWinGroupMembers->GridGroupMembers->Refresh();
        $$refWinGroupMembers->GridGroupMembers->BringWindowToTop();
      }
    }
  } else { Win32::GUI::MessageBox($$refWinGroupMembers, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010); }

}  #--- End loadDumpGroupMembers

#--------------------------#
sub loadDumpContacts
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinContacts, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $currAccountCode = $$refMech->selector('a._2s25', any => 1);
  my $currURL         = $$refMech->uri();
  if ($currURL =~ /https:\/\/(?:www|web).facebook.com\/messages/ and $currAccountCode->{href} =~ /\/([^\/]+)$/) {
    my $title = $1;
    if ($title =~ /profile.php\?id=([^\/\&]+)/) { $title = $1; }
    $title =~ s/[\#\<\>\:\"\/\\\|\?\*]/_/g;
    $$refWinContacts->tfContactsName->Text("$title - $$refSTR{'Contacts'}");
    $$refWinContacts->tfContactsCurrURL->Text($$refMech->uri);
  } else { Win32::GUI::MessageBox($$refWinContacts, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010); }

}  #--- End loadDumpContacts

#--------------------------#
sub loadDumpChat
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinChat, $USERDIR, $DEBUG_FILE, $refCONFIG, $refWin, $refSTR) = @_;
  my $currURL  = $$refMech->uri();
  if ($currURL !~ /https:\/\/(?:www|web).facebook.com\/messages\/t\/search\/([^\/\?]+)\/?/ and
      $currURL !~ /https:\/\/(?:www|web).facebook.com\/messages\/t\/([^\/\?]+)\/?/         and
      $currURL !~ /https:\/\/(?:www|web).facebook.com\/messages\/archived\/t\/([^\/\?]+)\/?/) {
    my $currTitle;
    ($currURL, $currTitle) = &loadPage($refMech, 'https:www.facebook.com/messages/t', $$refCONFIG{'TIME_TO_WAIT'})
    if $$refCONFIG{'AUTO_LOAD_SCROLL'};
  }
  if ($currURL =~ /https:\/\/(?:www|web).facebook.com\/messages\/t\/search\/([^\/\?]+)\/?/ or
      $currURL =~ /https:\/\/(?:www|web).facebook.com\/messages\/t\/([^\/\?]+)\/?/         or
      $currURL =~ /https:\/\/(?:www|web).facebook.com\/messages\/archived\/t\/([^\/\?]+)\/?/) {
    my $title = "$$refSTR{'Chat'} - $1";
    $title    =~ s/[\#\<\>\:\"\/\\\|\?\*]/_/g;
    my $currAccountCode = $$refMech->selector('a._2s25', any => 1);
    if ($currAccountCode->{href} =~ /\/([^\/]+)$/) {
      my $account = $1;
      $account =~ s/[\#\<\>\:\"\/\\\|\?\*]/_/g;
      if ($account =~ /id=(\d+)/) { $account = $1; }
      $title = "$account - $title";
    }
    $$refWinChat->tfChatName->Text($title);
    $$refWinChat->tfChatCurrURL->Text($currURL);
  } else { Win32::GUI::MessageBox($$refWinChat, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010); }

}  #--- End loadDumpChat

#--------------------------#
sub isDumpAlbumsReady
#--------------------------#
{
  # Local variables
  my $refWinAlbums = shift;
  my $saveDir      = $$refWinAlbums->tfDirSaveAlbums->Text();
  # No album title, no profile type or no valid directory
  if (!$$refWinAlbums->tfAlbumTitle->Text() or !$$refWinAlbums->tfPageType->Text() or !$saveDir or !-d $saveDir) {
    $$refWinAlbums->btnAlbumsDumpNow->Disable();
    $$refWinAlbums->btnAlbumsAddQueue->Disable();
    return(0);
  }
  # Selected report format
	my $selFormat = $$refWinAlbums->cbAlbumsFormat->GetCurSel();
	if ($selFormat == 1) {
		$$refWinAlbums->chAlbumsIncSmall->Enable();
		$$refWinAlbums->chAlbumsIncLarge->Enable();
		$$refWinAlbums->chAlbumsIncVideos->Enable();
	} else {
		$$refWinAlbums->chAlbumsIncSmall->Disable();
		$$refWinAlbums->chAlbumsIncLarge->Disable();
		$$refWinAlbums->chAlbumsIncVideos->Disable();
		$$refWinAlbums->chAlbumsIncSmall->Checked(0);
		$$refWinAlbums->chAlbumsIncLarge->Checked(0);
		$$refWinAlbums->chAlbumsIncVideos->Checked(0);
	}
  # Albums name loaded and at least one checked ?
  my $albumPicsChecked  = 0;
  my $albumVideoChecked = 0;
  for (my $i = 1; $i < $$refWinAlbums->GridAlbums->GetRows(); $i++) {
    if ($$refWinAlbums->GridAlbums->GetCellCheck($i, 0)) {
      if ($$refWinAlbums->GridAlbums->GetCellText($i, 2) =~ /vb\./ or $$refWinAlbums->GridAlbums->GetCellText($i, 2) =~ /videos/) {
        $albumVideoChecked = 1;
      } else { $albumPicsChecked = 1; }
      last if $albumPicsChecked and $albumVideoChecked;
    }
  }
  # Pic album selected
  if ($albumPicsChecked and $selFormat == 1) { $$refWinAlbums->chAlbumsIncLarge->Enable(); }
  else {
    $$refWinAlbums->chAlbumsIncLarge->Checked(0);
    $$refWinAlbums->chAlbumsIncLarge->Disable();
  }
  # Video album selected
  if ($albumVideoChecked and $selFormat == 1) { $$refWinAlbums->chAlbumsIncVideos->Enable(); }
  else {
    $$refWinAlbums->chAlbumsIncVideos->Checked(0);
    $$refWinAlbums->chAlbumsIncVideos->Disable();
  }
  # No selected albums
  if (!$albumPicsChecked and !$albumVideoChecked) {
    $$refWinAlbums->btnAlbumsDumpNow->Disable();
    $$refWinAlbums->btnAlbumsAddQueue->Disable();
    return(0);
  }
  $$refWinAlbums->btnAlbumsDumpNow->Enable();
  $$refWinAlbums->btnAlbumsAddQueue->Enable();

}  #--- End isDumpAlbumsReady

#--------------------------#
sub isDumpFriendsReady
#--------------------------#
{
  # Local variables
  my $refWinFriends = shift;
  my $saveDir       = $$refWinFriends->tfDirSaveFriends->Text();
  my $friendName    = $$refWinFriends->tfFriendName->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$friendName) { $$refWinFriends->btnFriendsDumpNow->Disable(); return(0); }
  # Friends category loaded and at least one checked ?
  my $friendsChecked = 0;
  for (my $i = 1; $i < $$refWinFriends->GridFriends->GetRows(); $i++) {
    if ($$refWinFriends->GridFriends->GetCellCheck($i, 0)) {
      $friendsChecked = 1;
      last;
    }
  }
  if (!$friendsChecked) {
    $$refWinFriends->btnFriendsDumpNow->Disable();
    $$refWinFriends->btnFriendsAddQueue->Disable();
    return(0);
  }
  $$refWinFriends->btnFriendsDumpNow->Enable();
  $$refWinFriends->btnFriendsAddQueue->Enable();

}  #--- End isDumpFriendsReady

#--------------------------#
sub isDumpMutualFriendsReady
#--------------------------#
{
  # Local variables
  my $refWinMutualFriends = shift;
  my $mutualFriendsName = $$refWinMutualFriends->tfMutualFriendsName->Text();
  my $saveDir      = $$refWinMutualFriends->tfDirSaveMutualFriends->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$mutualFriendsName) {
    $$refWinMutualFriends->btnMutualFriendsDumpNow->Disable();
    $$refWinMutualFriends->btnMutualFriendsAddQueue->Disable();
    return(0);
  }
  $$refWinMutualFriends->btnMutualFriendsDumpNow->Enable();
  $$refWinMutualFriends->btnMutualFriendsAddQueue->Enable();

}  #--- End isDumpMutualFriendsReady

#--------------------------#
sub isDumpContribReady
#--------------------------#
{
  # Local variables
  my $refWinContrib = shift;
  my $contribName   = $$refWinContrib->tfContribName->Text();
  my $saveDir       = $$refWinContrib->tfDirSaveContrib->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$contribName) {
    $$refWinContrib->btnContribDumpNow->Disable();
    $$refWinContrib->btnContribAddQueue->Disable();
    return(0);
  }
  # At least one type checked
  if (!$$refWinContrib->chContribComments->Checked()  and !$$refWinContrib->chContribLikes->Checked() and
      !$$refWinContrib->chContribVPosts->Checked()    and !$$refWinContrib->chContribEventPosts->Checked()) {
    $$refWinContrib->btnContribDumpNow->Disable();
    $$refWinContrib->btnContribAddQueue->Disable();
    return(0);
  }
  $$refWinContrib->btnContribDumpNow->Enable();
  $$refWinContrib->btnContribAddQueue->Enable();

}  #--- End isDumpContribReady

#--------------------------#
sub isDumpEventMembersReady
#--------------------------#
{
  # Local variables
  my $refWinEvent = shift;
  my $eventName   = $$refWinEvent->tfEventFilename->Text();
  my $saveDir     = $$refWinEvent->tfDirSaveEvent->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$eventName) {
    $$refWinEvent->btnEventDumpNow->Disable();
    $$refWinEvent->btnEventAddQueue->Disable();
    return(0);
  }
  # No selected lists
  if (!$$refWinEvent->chGoing->Checked()   and !$$refWinEvent->chMaybe->Checked()   and
      !$$refWinEvent->chInvited->Checked() and !$$refWinEvent->chDeclined->Checked()) {
    $$refWinEvent->btnEventDumpNow->Disable();
    $$refWinEvent->btnEventAddQueue->Disable();
    return(0);
  }
  $$refWinEvent->btnEventDumpNow->Enable();
  $$refWinEvent->btnEventAddQueue->Enable();

}  #--- End isDumpEventMembersReady

#--------------------------#
sub isDumpGroupMembersReady
#--------------------------#
{
  # Local variables
  my $refWinGroupMembers = shift;
  my $GroupMembersName   = $$refWinGroupMembers->tfGroupMembersName->Text();
  my $saveDir            = $$refWinGroupMembers->tfDirSaveGroupMembers->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$GroupMembersName) {
    $$refWinGroupMembers->btnGroupMembersDumpNow->Disable();
    $$refWinGroupMembers->btnGroupMembersAddQueue->Disable();  
    return(0);
  }
  # No selected lists
  my $nbrSel = 0;
  for (my $i = 1; $i < $$refWinGroupMembers->GridGroupMembers->GetRows(); $i++) {
    if ($$refWinGroupMembers->GridGroupMembers->GetCellCheck($i, 0)) {
      $nbrSel = 1;
      last;
    }
  }
  if (!$nbrSel) {
    $$refWinGroupMembers->btnGroupMembersDumpNow->Disable();
    $$refWinGroupMembers->btnGroupMembersAddQueue->Disable();
    return(0);
  }
  $$refWinGroupMembers->btnGroupMembersDumpNow->Enable();
  $$refWinGroupMembers->btnGroupMembersAddQueue->Enable();  

}  #--- End isDumpGroupMembersReady

#--------------------------#
sub isDumpContactsReady
#--------------------------#
{
  # Local variables
  my $refWinContacts = shift;
  my $contactsName = $$refWinContacts->tfContactsName->Text();
  my $saveDir      = $$refWinContacts->tfDirSaveContacts->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$contactsName) {
    $$refWinContacts->btnContactsDumpNow->Disable();
    $$refWinContacts->btnContactsAddQueue->Disable();
    return(0);
  }
  $$refWinContacts->btnContactsDumpNow->Enable();
  $$refWinContacts->btnContactsAddQueue->Enable();

}  #--- End isDumpContactsReady

#--------------------------#
sub isDumpChatReady
#--------------------------#
{
  # Local variables
  my $refWinChat = shift;
  my $chatName   = $$refWinChat->tfChatName->Text();
  my $saveDir    = $$refWinChat->tfDirSaveChat->Text();
  # Valid directory and valid name for save ?
  if (!$saveDir or !(-d $saveDir) or !$chatName) {
    $$refWinChat->btnChatDumpNow->Disable();
    $$refWinChat->btnChatAddQueue->Disable();
    return(0);
  }
  $$refWinChat->btnChatDumpNow->Enable() if $$refWinChat->rbChatCurrent->Checked();
  $$refWinChat->btnChatAddQueue->Enable();

}  #--- End isDumpChatReady

#--------------------------#
sub dumpAlbums
#--------------------------#
{
  # Local variables
  my ($now, $refWinAlbums, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
	&rememberPosWin($refWinAlbums, 'WINALBUMS', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpAlbums';
	$dumpParams{charSet}        = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinAlbums->tfAlbumTitle->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}   	    = encode($dumpParams{charSet}, $$refWinAlbums->tfDirSaveAlbums->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
	$dumpParams{startingURL}	  = $$refWinAlbums->tfAlbumCurrURL->Text();
	# Gather selected album names and urls
  my %albums;
	for (my $i = 1; $i < $$refWinAlbums->GridAlbums->GetRows(); $i++) {
		if ($$refWinAlbums->GridAlbums->GetCellCheck($i, 0)) {
			my $albumId = $$refWinAlbums->GridAlbums->GetCellText($i, 2);
      $albums{$albumId}{name} = $$refWinAlbums->GridAlbums->GetCellText($i, 1);
			$albums{$albumId}{url}  = $$refWinAlbums->GridAlbums->GetCellText($i, 3);
		}
	}
	$dumpParams{pageType}				= $$refWinAlbums->tfPageType->Text(); # Page type: 0=unknown, 1=People, 2=Groups, 3=Pages (Business)
	$dumpParams{openAlbumDir}		= $$refWinAlbums->chAlbumsOpenDir->Checked();
	$dumpParams{incPublishDate} = $$refWinAlbums->chPublishDate->Checked();
	$dumpParams{incSmallPics}		= $$refWinAlbums->chAlbumsIncSmall->Checked();
	$dumpParams{incLargePics}		= $$refWinAlbums->chAlbumsIncLarge->Checked();
	$dumpParams{incVideos}      = $$refWinAlbums->chAlbumsIncVideos->Checked();
	$dumpParams{reportFormat}   = $$refWinAlbums->cbAlbumsFormat->GetString($$refWinAlbums->cbAlbumsFormat->GetCurSel());
  mkdir("$USERDIR\\Queue")    if !-d "$USERDIR\\Queue";
	if (&createDumpDB(  "$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams) and
      &createAlbumsDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%albums    )) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      if (Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR)) {
        &winAlbums_Terminate();
        $processObj->Wait(INFINITE);
        # Final message
        $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'Albums'} $$refSTR{'finished'}");
        if (!$$refWin->IsVisible()) {
          $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                                 -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'Albums'} $$refSTR{'finished'}");
          $$refWin->Tray->ShowBalloon(1);
        }
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winAlbums_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}

}  #--- End dumpAlbums

#--------------------------#
sub dumpFriends
#--------------------------#
{
  # Local variables
  my ($now, $refWinFriends, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
	&rememberPosWin($refWinFriends, 'WINFRIENDS', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpFriends';
	$dumpParams{charSet}        = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinFriends->tfFriendName->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}   	    = encode($dumpParams{charSet}, $$refWinFriends->tfDirSaveFriends->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
  $dumpParams{startingURL}	  = $$refWinFriends->tfFriendCurrURL->Text();
	# Gather selected friend categories names and urls
	for (my $i = 1; $i < $$refWinFriends->GridFriends->GetRows(); $i++) {
		if ($$refWinFriends->GridFriends->GetCellCheck($i, 0)) {
			my $catName = $$refWinFriends->GridFriends->GetCellText($i, 1);
			my $catid   = $$refWinFriends->GridFriends->GetCellText($i, 2);
			my $catURL  = $$refWinFriends->GridFriends->GetCellText($i, 3);
      $dumpParams{listCat}       .= $catName . '|';
			$dumpParams{"$catName-id"}  = $catid;
			$dumpParams{"$catName-url"} = $catURL;
		}
	}
  chop($dumpParams{listCat});
	$dumpParams{incIcons}			= $$refWinFriends->chFriendsProfileIcons->Checked();
	$dumpParams{reportFormat}	= $$refWinFriends->cbFriendsFormat->GetString($$refWinFriends->cbFriendsFormat->GetCurSel());
  mkdir("$USERDIR\\Queue")  if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winFriends_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'friends'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'friends'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winFriends_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}
  
}  #--- End dumpFriends

#--------------------------#
sub dumpMutualFriends
#--------------------------#
{
  # Local variables
  my ($now, $refWinMutualFriends, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR,
      $USERDIR, $refWin, $refSTR) = @_;
	&rememberPosWin($refWinMutualFriends, 'WINMUTUALFRIENDS', $refWinConfig, $refCONFIG, $CONFIG_FILE)
  if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpMutualFriends';
	$dumpParams{charSet}    	  = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinMutualFriends->tfMutualFriendsName->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}  	    = encode($dumpParams{charSet}, $$refWinMutualFriends->tfDirSaveMutualFriends->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
	$dumpParams{startingURL}	  = $$refWinMutualFriends->tfMutualFriendsCurrURL->Text();
	$dumpParams{incIcons}			  = 1 if $$refWinMutualFriends->chMutualFriendsProfileIcons->Checked();
	$dumpParams{autoScroll}     = 1 if $$refWinMutualFriends->chMutualFriendsAutoScroll->Checked();
	$dumpParams{reportFormat}   = $$refWinMutualFriends->cbMutualFriendsFormat->GetString($$refWinMutualFriends->cbMutualFriendsFormat->GetCurSel());
  mkdir("$USERDIR\\Queue")    if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winMutualFriends_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'MutualFriends'} $$refSTR{'finished'}...");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'MutualFriends'} $$refSTR{'finished'}...");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winMutualFriends_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}

}  #--- End dumpMutualFriends

#--------------------------#
sub dumpContrib
#--------------------------#
{
  # Local variables
  my ($now, $refWinContrib, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
  &rememberPosWin($refWinContrib, 'WINCONTRIB', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpContrib';
	$dumpParams{charSet}        = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinContrib->tfContribName->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}   	    = encode($dumpParams{charSet}, $$refWinContrib->tfDirSaveContrib->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
  $dumpParams{startingID}     = $$refWinContrib->tfContribID->Text() if $$refWinContrib->tfContribID->Text();
  $dumpParams{startingURL}	  = $$refWinContrib->tfStartURL->Text();
  $dumpParams{reportFormat}	  = $$refWinContrib->cbContribFormat->GetString($$refWinContrib->cbContribFormat->GetCurSel());
	$dumpParams{incIcons}			  = 1 if $$refWinContrib->chContribProfileIcons->Checked();
	$dumpParams{comments}			  = 1 if $$refWinContrib->chContribComments->Checked();
	$dumpParams{likes}				  = 1 if $$refWinContrib->chContribLikes->Checked();
	$dumpParams{vPosts}				  = 1 if $$refWinContrib->chContribVPosts->Checked();
	$dumpParams{eventPosts}     = 1 if $$refWinContrib->chContribEventPosts->Checked();
	$dumpParams{autoScroll}     = 1 if $$refWinContrib->chContribAutoScroll->Checked();
	# Create database
  mkdir("$USERDIR\\Queue") if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winContrib_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'contributors'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'contributors'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winContrib_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}
	
}  #--- End dumpContrib

#--------------------------#
sub dumpEventMembers
#--------------------------#
{
  # Local variables
  my ($now, $refWinEvent, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
  &rememberPosWin($refWinEvent, 'WINEVENT', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpEvent';
	$dumpParams{charSet}        = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinEvent->tfEventFilename->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}   	    = encode($dumpParams{charSet}, $$refWinEvent->tfDirSaveEvent->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
	$dumpParams{startingURL}	  = $$refWinEvent->tfEventCurrURL->Text();
	$dumpParams{incIcons}			  = 1 if $$refWinEvent->chEventProfileIcons->Checked();
	$dumpParams{DataURL}			  = $$refWinEvent->tfDataURL->Text();
	$dumpParams{reportFormat}	  = $$refWinEvent->cbEventFormat->GetString($$refWinEvent->cbEventFormat->GetCurSel());
	$dumpParams{going}			    = 1 if $$refWinEvent->chGoing->Checked();
	$dumpParams{maybe}			    = 1 if $$refWinEvent->chMaybe->Checked();
	$dumpParams{invited}		    = 1 if $$refWinEvent->chInvited->Checked();
	$dumpParams{declined}	      = 1 if $$refWinEvent->chDeclined->Checked();
	# Create database
  mkdir("$USERDIR\\Queue") if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winEvent_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'EventMembers'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'EventMembers'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winEvent_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}
  
}  #--- End dumpEventMembers

#--------------------------#
sub dumpGroupMembers
#--------------------------#
{
  # Local variables
  my ($now, $refWinGroupMembers, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
  &rememberPosWin($refWinGroupMembers, 'WINGROUP_MEMBERS', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpGroupMembers';
	$dumpParams{charSet}        = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinGroupMembers->tfGroupMembersName->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}   	    = encode($dumpParams{charSet}, $$refWinGroupMembers->tfDirSaveGroupMembers->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
  $dumpParams{startingURL}	  = $$refWinGroupMembers->tfGroupMembersCurrURL->Text();
	# Gather selected Group members categories names and urls
	for (my $i = 1; $i < $$refWinGroupMembers->GridGroupMembers->GetRows(); $i++) {
		if ($$refWinGroupMembers->GridGroupMembers->GetCellCheck($i, 0)) {
			my $catName = decode($$refCONFIG{'CHARSET'}, $$refWinGroupMembers->GridGroupMembers->GetCellText($i, 1));
      $dumpParams{listCat}       .= $catName . '|';
			$dumpParams{"$catName-url"} = $$refWinGroupMembers->GridGroupMembers->GetCellText($i, 2);
		}
	}
  chop($dumpParams{listCat});
  $dumpParams{groupType}		= $$refWinGroupMembers->tfGroupMembersType->Text();
	$dumpParams{incIcons}			= $$refWinGroupMembers->chGroupMembersProfileIcons->Checked();
	$dumpParams{reportFormat}	= $$refWinGroupMembers->cbGroupMembersFormat->GetString($$refWinGroupMembers->cbGroupMembersFormat->GetCurSel());
  mkdir("$USERDIR\\Queue")  if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winGroupMembers_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'groupMembers'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'groupMembers'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winGroupMembers_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}", $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}
	
}  #--- End dumpGroupMembers

#--------------------------#
sub dumpContacts
#--------------------------#
{
  # Local variables
  my ($now, $refWinContacts, $refWinQueue, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
	&rememberPosWin($refWinContacts, 'WINCONTACTS', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{procID}				  = time;
	$dumpParams{processName}	  = 'DumpContacts';
	$dumpParams{charSet}    	  = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{filename} 	    = encode($dumpParams{charSet}, $$refWinContacts->tfContactsName->Text());
	$dumpParams{filename} 	    =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
	$dumpParams{saveDir}  	    = encode($dumpParams{charSet}, $$refWinContacts->tfDirSaveContacts->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
	$dumpParams{startingURL}	  = $$refWinContacts->tfContactsCurrURL->Text();
	$dumpParams{incIcons}			  = 1 if $$refWinContacts->chContactsProfileIcons->Checked();
	$dumpParams{autoScroll}     = 1 if $$refWinContacts->chContactsAutoScroll->Checked();
	$dumpParams{reportFormat}   = $$refWinContacts->cbContactsFormat->GetString($$refWinContacts->cbContactsFormat->GetCurSel());
  mkdir("$USERDIR\\Queue")    if !-d "$USERDIR\\Queue";
	if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
    if ($now) { # Dump Now
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      &winContacts_Terminate();
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'Contacts'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'Contacts'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    } else { # Add to queue
      &createWinQueue() if !$$refWinQueue;
      if (&existsInQueue($refWinQueue, $dumpParams{filename})) {
        my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'queueExists'} ?", $$refSTR{'Queue'}, 0x40024);
        if ($answer == 7) { # Answer is no, abort, but keep window open
          unlink("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db");
          return(1);
        }
      }
      &winContacts_Terminate();
      if (&addToQueue($refWinQueue, "$dumpParams{processName}-$dumpParams{procID}",
                      $dumpParams{filename}, $dumpParams{startingURL})) {
        Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
      } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
    }
	}

}  #--- End dumpContacts

#--------------------------#
sub dumpChat
#--------------------------#
{
  # Local variables
  my ($now, $refTHR, $refWinChat, $refWinQueue, $refWinPb2, $refWinConfig, $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $refWin, $refSTR) = @_;
	&rememberPosWin($refWinChat, 'WINCHAT', $refWinConfig, $refCONFIG, $CONFIG_FILE) if $$refWinConfig->chRememberPos->Checked();
	# Get Dump parameters
	my %dumpParams;
	$dumpParams{processName}	  = 'DumpChat';
	$dumpParams{charSet}    	  = $$refWinConfig->cbCharset->GetString($$refWinConfig->cbCharset->GetCurSel());
	$dumpParams{saveDir}  	    = encode($dumpParams{charSet}, $$refWinChat->tfDirSaveChat->Text());
	chop($dumpParams{saveDir})  if $dumpParams{saveDir} =~ /\\$/;
	$dumpParams{debugLogging}   = 1 if $$refWinConfig->chDebugLogging->Checked();
	$dumpParams{timeToWait}     = $$refWinConfig->tfTimeToWait->Text();
	$dumpParams{silentProgress} = 1 if $$refWinConfig->chSilentProgress->Checked() and !$now;
	$dumpParams{closeUsedTabs}  = 1 if $$refWinConfig->chCloseUsedTabs->Checked();
	$dumpParams{delTempFiles}   = 1 if $$refWinConfig->chDelTempFiles->Checked();
  $dumpParams{openReport}		  = 1 if $$refWinConfig->chOptOpenReport->Checked() and ($now or !$$refWinConfig->chOptDontOpenReport->Checked());
	$dumpParams{autoScroll}     = 1 if $$refWinChat->chChatAutoScroll->Checked();
	$dumpParams{dlImages} 	    = 1 if $$refWinChat->chDownloadImg->Checked();
	$dumpParams{dlPictures} 	  = 1 if $$refWinChat->chDownloadPics->Checked();
	$dumpParams{dlAttached} 	  = 1 if $$refWinChat->chDownloadAD->Checked();
	$dumpParams{dlVideos}   	  = 1 if $$refWinChat->chDownloadVid->Checked();
	$dumpParams{dlVocalMsg}     = 1 if $$refWinChat->chDownloadVM->Checked();
	$dumpParams{dateRange}  	  = $$refWinChat->rbChatDatesRange->Checked();
	if ($dumpParams{dateRange}) {
		my ($d1, $m1, $y1)      = $$refWinChat->dtChatDatesRangeS->GetDate();
		$dumpParams{dateStart}  = timelocal(0,0,0,$d1,$m1-1,$y1); # Store in Unixtime format
		my ($d2, $m2, $y2)      = $$refWinChat->dtChatDatesRangeE->GetDate();
		$dumpParams{dateEnd}    = timelocal(0,0,0,$d2,$m2-1,$y2) + 86399; # Store in Unixtime format
	}
	$dumpParams{reportFormat} = $$refWinChat->cbChatFormat->GetString($$refWinChat->cbChatFormat->GetCurSel());
  # Scroll options
  $dumpParams{maxScrollChatByDate} = $$refWinConfig->rbMaxScrollChatByDate->Checked();
  if ($dumpParams{maxScrollChatByDate}) {
    my ($d, $m, $y)      = $$refWinConfig->dtMaxScrollChatByDate->GetDate();
    $dumpParams{maxDate} = timelocal(0,0,0,$d,$m-1,$y); # Store in Unixtime format
  } else { $dumpParams{maxScrollChat} = $$refWinConfig->tfMaxScrollChat->Text(); }
  mkdir("$USERDIR\\Queue") if !-d "$USERDIR\\Queue";
  if ($now or $$refWinChat->rbChatCurrent->Checked()) {
    $dumpParams{procID}      = time;
    $dumpParams{filename} 	 = encode($dumpParams{charSet}, $$refWinChat->tfChatName->Text());
    $dumpParams{filename} 	 =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
    $dumpParams{startingURL} = $$refWinChat->tfChatCurrURL->Text();
  }
  # Dump Now
  &winChat_Terminate();
  if ($now) {
    if (&createDumpDB("$USERDIR\\Queue\\$dumpParams{processName}-$dumpParams{procID}\.db", \%dumpParams)) {
      my $command = 'ExtractFace-process ' . "$dumpParams{processName} $dumpParams{procID} \"$PROGDIR\" \"$USERDIR\"";
      Win32::Process::Create(my $processObj, $PROGDIR .'\ExtractFace-process.exe', $command, 0, NORMAL_PRIORITY_CLASS, $PROGDIR);
      $processObj->Wait(INFINITE);
      # Final message
      $$refWin->Tray->Change(-tip => "$$refSTR{'Dump'} $$refSTR{'Chat'} $$refSTR{'finished'}");
      if (!$$refWin->IsVisible()) {
        $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                               -balloon_tip => "$$refSTR{'Dump'} $$refSTR{'Chat'} $$refSTR{'finished'}");
        $$refWin->Tray->ShowBalloon(1);
      }
    }
  # Add to queue
  } else {
    &createWinQueue() if !$$refWinQueue;
    my $return  = 0;
    my $nbrChat = 0;
    # Current chat only
    if ($$refWinChat->rbChatCurrent->Checked()) {
      $dumpParams{procID}      = time;
      $dumpParams{filename} 	 = encode($dumpParams{charSet}, $$refWinChat->tfChatName->Text());
      $dumpParams{filename} 	 =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
      $dumpParams{startingURL} = $$refWinChat->tfChatCurrURL->Text();
      $return  = &addDumpChatToQueue(\%dumpParams, $refWinQueue, $USERDIR);
      $nbrChat = 1;
    # All selected chat must be added to queue
    } else {
      # Avoid new thread
      if ($$refTHR and $$refTHR->is_running()) { Win32::GUI::MessageBox($$refWin, $$refSTR{'processRunning'},$$refSTR{'Warning'},0x40010); }
      else {
        $$refTHR = threads->create(sub {
          # Cancel button
          $SIG{'KILL'} = sub {
            # Turn off progress bar
            $$refWinPb2->lblPbCurr->Text('');
            $$refWinPb2->lblCount->Text('');
            &winPb2_Terminate;
            threads->exit();
          };
          $SIG{__DIE__} = sub {
            my $err = (split(/ at /, $_[0]))[0];
            Win32::GUI::MessageBox($$refWinChat, "$$refSTR{'processCrash'}: $err", $$refSTR{'Error'}, 0x40010);
          };
          # Turn on progress bar
          for (my $i = 1; $i < $$refWinChat->GridChats->GetRows(); $i++) {
            $nbrChat++ if $$refWinChat->GridChats->GetCellCheck($i, 0);
          }
          $$refWinPb2->lblLogo->Show();
          $$refWinPb2->btnCancel2->Show();
          $$refWinPb2->Center($$refWin);
          $$refWinPb2->Show();
          $$refWinPb2->lblPbCurr->Text($$refSTR{'addingQueue'}.'...');
          $$refWinPb2->lblCount->Text("0/$nbrChat");
          my $count = 0;
          for (my $i = 1; $i < $$refWinChat->GridChats->GetRows(); $i++) {
            if ($$refWinChat->GridChats->GetCellCheck($i, 0)) {
              $dumpParams{procID}      = gettimeofday;
              $dumpParams{procID}      =~ s/\.//;
              $dumpParams{filename} 	 = $$refWinChat->GridChats->GetCellText($i, 2);
              $dumpParams{filename} 	 =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g; # Remove invalid characters for Windows filename
              $dumpParams{startingURL} = $$refWinChat->GridChats->GetCellText($i, 3);
              $return += &addDumpChatToQueue(\%dumpParams, $refWinQueue, $USERDIR);
              $count++;
              $$refWinPb2->lblCount->Text("$count/$nbrChat");
            }
          }
          # Turn off progress bar
          $$refWinPb2->lblPbCurr->Text('');
          $$refWinPb2->lblCount->Text('');
          &winPb2_Terminate;
          if ($return == $nbrChat) {
            Win32::GUI::MessageBox($$refWin, $$refSTR{'addedQueue'}.'!', $$refSTR{'Queue'}, 0x40040);
          } else { Win32::GUI::MessageBox($$refWin, $$refSTR{'errAddQueue'}, $$refSTR{'Error'}, 0x40010); }
        });
      }
    }
  }

}  #--- End dumpChat

#--------------------------#
sub validAlbumPage
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinAlbums, $refCONFIG, $refSTR) = @_;
  my ($pageType, $goodAlbumUrl, $currTitle);
  my $currURL = $$refMech->uri();
  my $valid   = 0;
  if ($currURL !~ /photos_albums/        and $currURL !~ /photos\?lst=[\w\%]+\&collection_token=\w+%\w+%3A6/ and
			$currURL !~ /photos\/\?tab=albums/ and $currURL !~ /sk=photos\&collection_token=\w+%\w+%3A6/ and
      $currURL !~ /photos\/\?filter=albums/) { # Not in the album page
    return(0) if !$$refCONFIG{'AUTO_LOAD_SCROLL'}; # Don't load and scroll automatically
    # Determine type of page
    ($pageType, $goodAlbumUrl) = &guessPageType($refMech, $currURL); # Page type: 0=unknown, 1=People, 2=Groups, 3=Pages (Business)
    # Load the good Album page
    ($currURL, $currTitle) = &loadPage($refMech, $goodAlbumUrl, $$refCONFIG{'TIME_TO_WAIT'}) if $pageType and $goodAlbumUrl;
    $$refWinAlbums->tfPageType->Text($pageType) if $pageType;
    # Re evaluate current page
    if (($currURL !~ /photos_albums/        and $currURL !~ /photos_albums\?/         and $currURL !~ /collection_token=\w+%\w+%3A6/) and
				 $currURL !~ /photos\/\?tab=albums/ and $currURL !~ /photos\/\?filter=albums/ or  $currTitle =~ /Page Not Found/) {  # Still not in the right page
      $$refWinAlbums->btnAlbumsDumpNow->Disable();
      $$refWinAlbums->btnAlbumsAddQueue->Disable();
      Win32::GUI::MessageBox($$refWinAlbums, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010);
      threads->exit();
    } else { # Now in the right page, scroll down to load the whole page
      $$refWinAlbums->lblInProgress->Text($$refSTR{'loadAlbum'}.'...');
      if ($pageType == 1) { &scrollAlbumPage($refMech, $$refCONFIG{'TIME_TO_WAIT'}); }
      else {
        my $end = &scrollPage($refMech, $$refCONFIG{'TIME_TO_WAIT'});
        while (!$end) { $end = &scrollPage($refMech, $$refCONFIG{'TIME_TO_WAIT'}); }
      }
      $$refMech->eval_in_page('window.scrollTo(0,0)') if $$refCONFIG{'OPT_SCROLL_TOP'};
      $valid = 1;
    }
  # You are in the right page, scroll down to load the whole page
  } else {
    # Determine type of page
    ($pageType, $goodAlbumUrl) = &guessPageType($refMech, $currURL); # Page type: 0=unknown, 1=People, 2=Groups, 3=Pages (Business)
    $$refWinAlbums->tfPageType->Text($pageType) if $pageType;
    # Get current page title
    if ($currURL and $currURL =~ /https:\/\/(?:www|web).facebook.com\//) {
      if    ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/) { $currTitle = $1; }
			elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/groups\/([^\/\?]+)/        ) { $currTitle = $1; }
      elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/                ) { $currTitle = $1; }
    }
    $$refWinAlbums->lblInProgress->Text($$refSTR{'loadAlbum'}.'...');
    if ($$refCONFIG{'AUTO_LOAD_SCROLL'}) { # Scroll the page
      if ($pageType == 1) { &scrollAlbumPage($refMech, $$refCONFIG{'TIME_TO_WAIT'}); }
      else {
        my $end = &scrollPage($refMech, $$refCONFIG{'TIME_TO_WAIT'});
        while (!$end) { $end = &scrollPage($refMech, $$refCONFIG{'TIME_TO_WAIT'}); }
      }
    }
    $$refMech->eval_in_page('window.scrollTo(0,0)')    if $$refCONFIG{'OPT_SCROLL_TOP'};
    $valid = 1;
  }
  # Write title
  if ($valid and $currTitle) {
    chop($currTitle) if $currTitle =~ /#$/;
    $currTitle .= " - Albums";
    $$refWinAlbums->tfAlbumTitle->Text($currTitle);
  }
  return($valid, $currURL);

}  #--- End validAlbumPage

#--------------------------#
sub getListAlbums
#--------------------------#
{
  # Local variables
  my ($refMech, $refWinAlbums, $pageType, $USERDIR, $refCONFIG) = @_;
  my $tempDir = "$USERDIR\\temp";
  mkdir("$tempDir") if !-d "$tempDir";
  my $htmlPage = "$tempDir\\temp.html";
  my $status = $$refMech->save_content($htmlPage);
  while ($status->{currentState} != $status->{PERSIST_STATE_FINISHED}) { sleep(1); }
  # Load content from file
  if (-T $htmlPage and open(my $fhTemp, "<:encoding(utf8)", $htmlPage)) {
    my $file_as_string = do { local $/ = <$fhTemp> };
    $file_as_string =~ s/[\r\n]//g;
    close($fhTemp);
    my @albumsCode;
    if    ($pageType == 1) { @albumsCode = split(/_51m-/         , $file_as_string); }
    elsif ($pageType == 3) { @albumsCode = split(/_3rte/         , $file_as_string); }
    else                   { @albumsCode = split(/photoTextTitle/, $file_as_string); }
    shift(@albumsCode);
    my %tmpAlbums;
    foreach my $albumCode (@albumsCode) {
      my $url;
      my $name;
      my $id;
      if (($pageType == 1 or $pageType == 3) and $albumCode =~ /href="([^\"]+)"/) {
        $url = $1;
        if    ($albumCode =~ /_50f4[^\>]*>([^\<]*)</) { $name = $1; }
        if    ($url       =~ /album_id=(\d+)/       ) { $id   = $1; }
        elsif ($url       =~ /set=([^\&]+)/         ) { $id   = $1; }
      } elsif ($albumCode =~ /href="([^\"]+)"><strong>([^\<]*)\</ or
               $albumCode =~ /href="([^\"]+)"><i[^\>]+><\/i><strong>([^\<]*)\</) {
        $url  = $1;
        $name = $2;
        if ($url =~ /set=([^\&]+)/) { $id = $1; }
      }
      if ($url and $name and $id) {
        $url =~ s/&amp;/&/g;
        $tmpAlbums{$id}{name} = encode($$refCONFIG{'CHARSET'}, $name);
        $tmpAlbums{$id}{url}  = $url;
      }
    }
    # Feed the grid
    foreach my $id (keys %tmpAlbums) {
      if (my $i = $$refWinAlbums->GridAlbums->InsertRow($tmpAlbums{$id}{name}, -1)) {
        $$refWinAlbums->GridAlbums->SetCellText($i, 0, ''        );
        $$refWinAlbums->GridAlbums->SetCellType($i, 0, GVIT_CHECK);
        $$refWinAlbums->GridAlbums->SetCellCheck($i, 0, 1);
        $$refWinAlbums->GridAlbums->SetCellText($i, 1, $tmpAlbums{$id}{name});
        $$refWinAlbums->GridAlbums->SetCellText($i, 2, $id                  );
        $$refWinAlbums->GridAlbums->SetCellText($i, 3, $tmpAlbums{$id}{url} );
        $$refWinAlbums->GridAlbums->Refresh();
      }
    }
  }
  
}  #--- End getListAlbums
  
#--------------------------#
sub guessPageType
#--------------------------#
{
  # Local variables
  my ($refMech, $currURL) = @_;
  my $pageType = 0; # 0 = unknown, 1 = People, 2 = Groups, 3 = Pages (Business)
  my $goodAlbumUrl;
  # Trying to get the good page
  if      ($currURL =~ /https:\/\/(?:www|web).facebook.com\/groups\/([^\/]+)/ ) { # Album from a Group page (public)
    $goodAlbumUrl = "https://www.facebook.com/groups/$1/photos/?filter=albums";
    $pageType     = 2;
  } elsif  ($currURL =~ /https:\/\/(?:www|web).facebook.com\/pg\/([^\/]+)/    ) { # Page (Business)
    $goodAlbumUrl = "https://www.facebook.com/$1/photos/?tab=albums";
    $pageType     = 3;
  } elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/ ) { # People: Profile with id
    $goodAlbumUrl = "https://www.facebook.com/profile.php?id=$1&sk=photos&collection_token=$1%3A2305272732%3A6";
    $pageType     = 1;
  } elsif ($currURL =~ /(?:\?lst=\d+%3A\d+%3A\d+|\?collection_token=\d+)/ and
                         $currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\?\/]+)/) { # People: Profile with id
    $goodAlbumUrl = "https://www.facebook.com/$1/photos_albums";
    $pageType     = 1;
  } elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/]+)/) { # People - Other or Pages
    my $searchTag = $$refMech->selector('a._2wmb', any => 1); # The search tag is specific to Pages (ex.: @businessName)
    if ($searchTag) { $goodAlbumUrl = "https://www.facebook.com/$1/photos/?tab=albums"; $pageType = 3; } # Page (Business)
    else            { $goodAlbumUrl = "https://www.facebook.com/$1/photos_albums";      $pageType = 1; } # People
  }
  return($pageType, $goodAlbumUrl);
  
}  #--- End guessPageType

#--------------------------#
sub scrollAlbumPage
#--------------------------#
{
  # Local variables
  my ($refMech, $time) = @_;
  while (1) {
    sleep($time); # End of the page ?
    my @parts = $$refMech->selector('div._30f');
    if (scalar(@parts) > 1) { return(1); } # End of the page
    else { # Scrolling down and wait for content to load
      $$refMech->eval_in_page('window.scrollTo(0,document.body.scrollHeight)');
      sleep($time);
    }
  }

}  #--- End scrollAlbumPage

#--------------------------#
sub handlePageThr
#--------------------------#
{
  # Local variables
  my ($refTHR, $typeHandle, $nbrRetries, $count, $refARROW, $refHOURGLASS, $DEBUG_FILE, $refCONFIG, $refWinConfig,
      $refWinPb2, $refWin, $refSTR) = @_;
  # $typeHandle: 1=Scroll, 2=Expand, 3=Scroll and Expand, 4=Scroll contacts, 5=Scroll chat, 6=Load Newer Msg, 7=Load Older Msg, 8=Remove header (blue bar, menu),
  #              9=Remove left column, 10=Remove rigth column, 11=Remove bottom, 12=Remove all
  my @processName = (undef, $$refSTR{'Scrolling'}, $$refSTR{'Expanding'}, $$refSTR{'ScrollExpand'}, "$$refSTR{'Scrolling'} $$refSTR{'Contacts'}",
                     "$$refSTR{'Scrolling'} $$refSTR{'Chat'}", $$refSTR{'loadNewerMsg'}, $$refSTR{'loadOlderMsg'}, "$$refSTR{'Remove'} $$refSTR{'Top'}",
                     "$$refSTR{'Remove'} $$refSTR{'leftCol'}", "$$refSTR{'Remove'} $$refSTR{'rightCol'}", "$$refSTR{'Remove'} $$refSTR{'Bottom'}",
                     "$$refSTR{'Remove'} $$refSTR{'All'}");
  # Cancel button
  $SIG{'KILL'} = sub {
    # Turn off progress bar
    $$refWinPb2->lblPbCurr->Text('');
    $$refWinPb2->lblCount->Text('');
    &winPb2_Terminate;
    $$refWin->ChangeCursor($$refARROW);
    $$refWin->Tray->Change(-tip => "$processName[$typeHandle] $$refSTR{'cancelled'}");
    threads->exit();
  };
  # Deal with crash
  $SIG{__DIE__} = sub {
    my $msgErr = $_[0];
    chomp($msgErr);
    $msgErr =~ s/[\t\r\n]/ /g;
    &debug($msgErr, $DEBUG_FILE) if $$refCONFIG{'DEBUG_LOGGING'};
    if ($msgErr =~ /problem connecting/) {
      Win32::GUI::MessageBox($$refWin, $$refSTR{'errMozRepl'}, $$refSTR{'Error'}, 0x40010);
      # Turn off progress bar
      $$refWinPb2->lblPbCurr->Text('');
      $$refWinPb2->lblCount->Text('');
      &winPb2_Terminate;
      $$refWin->ChangeCursor($$refARROW);
      $$refWin->Tray->Change(-tip => 'ExtractFace');
    } else {
      # Retry 10 times
      $nbrRetries++ if $msgErr !~ /NS_ERROR_FILE_IS_LOCKED/;
      if ($nbrRetries < $$refCONFIG{'NBR_RESUME'}) {
        # Restart a new thread to continue
        $$refWinPb2->lblPbCurr->Text($$refSTR{'crash'}.'...') if $msgErr !~ /NS_ERROR_FILE_IS_LOCKED/;
        sleep(2);
        $$refTHR = threads->create(\&handlePageThr, $refTHR, $typeHandle, $nbrRetries, $count, $refARROW, $refHOURGLASS,
                                   $DEBUG_FILE, $refCONFIG, $refWinConfig, $refWinPb2, $refWin, $refSTR);
      } else {
				my $err = (split(/ at /, $msgErr))[0];
				Win32::GUI::MessageBox($$refWin, "$$refSTR{'processCrash'}: $err", $$refSTR{'Error'}, 0x40010);
        # Turn off progress bar
        $$refWinPb2->lblPbCurr->Text('');
        $$refWinPb2->lblCount->Text('');
        &winPb2_Terminate;
        $$refWin->ChangeCursor($$refARROW);
        $$refWin->Tray->Change(-tip => 'ExtractFace');
      }
    }
    threads->exit();
  };
  # First execution
  if (!$nbrRetries) {
    $$refWin->ChangeCursor($$refHOURGLASS);
    $$refWin->Tray->Change(-tip => "$processName[$typeHandle] $$refSTR{'inProgress'}...");
    # Turn on progress bar
    $$refWinPb2->Center($$refWin);
    $$refWinPb2->Show();
    $$refWinPb2->lblPbCurr->Text('');
    $$refWinPb2->lblCount->Text('');
    $$refWin->Disable();
  }
  my $mech;
  eval { $mech = WWW::Mechanize::Firefox->new(tab => 'current'); };
  if ($@) {
    Win32::GUI::MessageBox($$refWin, $$refSTR{'errMozRepl'}, $$refSTR{'Error'}, 0x40010)
    if $@ =~ /Failed to connect to/;
    threads->exit();
  }
  if ($mech->uri() =~ /facebook.com/) {
    $$refWinPb2->lblPbCurr->Text("$processName[$typeHandle] $$refSTR{'inProgress'}...");
    # Scroll
    if      ($typeHandle == 1) {
      &scrollToBottom(\$mech, $$refCONFIG{'DEBUG_LOGGING'}, $count, $refWinConfig);
      $mech->eval_in_page('window.scrollTo(0,0)') if $$refCONFIG{'OPT_SCROLL_TOP'}; # Scroll to the top
    # Expand
    } elsif ($typeHandle == 2) {
      for (0..2) { &expandContent(\$mech, $refCONFIG); }
    # Scroll and Expand
    } elsif ($typeHandle == 3) {
      my $maxScrollByDate = $$refWinConfig->rbMaxScrollByDate->Checked();
      my $maxDate;
      my $maxScroll;
      if ($maxScrollByDate) {
        my ($d, $m, $y) = $$refWinConfig->dtMaxScrollByDate->GetDate();
        $maxDate        = timelocal(0,0,0,$d,$m-1,$y); # Store in Unixtime format
      } else { $maxScroll = $$refWinConfig->tfMaxScroll->Text(); }
        my $end = &scrollPage(\$mech, $$refCONFIG{'TIME_TO_WAIT'});
        $count++;
        while (!$end) { # If $end == 1, we reached the end of the page
          # Expand
          if ($$refCONFIG{'EXPAND_SEE_MORE'} or $$refCONFIG{'EXPAND_MORE_POSTS'} or $$refCONFIG{'SEE_TRANSLATION'}) {
            # Expand all additional content
            $$refWinPb2->lblPbCurr->Text("$$refSTR{'Expanding'} $$refSTR{'inProgress'}...");
            $$refWin->Tray->Change(-tip => "$$refSTR{'Expanding'} $$refSTR{'inProgress'}...");
            &expandContent(\$mech, $refCONFIG);
            &expandContent(\$mech, $refCONFIG); # Do it again
          }
          # Scroll again
          $$refWinPb2->lblPbCurr->Text("$$refSTR{'Scrolling'} $$refSTR{'inProgress'}...");
          if ($maxScrollByDate and $maxDate) { # Stop by date
            my $lastDisplayedDate = ($mech->selector('a._5pcq abbr'))[-1];
            if ($lastDisplayedDate->{outerHTML} =~ /data-utime="([^\"]+)"/) {
              my $date = $1;
              last if $date <= $maxDate;
            }
          } elsif ($maxScroll and $count >= $maxScroll) { last; } # Stop by page
          $end = &scrollPage(\$mech, $$refCONFIG{'TIME_TO_WAIT'});
          $count++;
        }
        # Scroll to the top
        $mech->eval_in_page('window.scrollTo(0,0)') if $$refCONFIG{'OPT_SCROLL_TOP'};
      } elsif ($typeHandle == 4) { # Scroll contacts
        while (1) {
          if (my $loadMoreContacts = ($mech->selector('div._19hf a', any => 1))[0]) {
            $mech->eval_in_page("var scrollingDiv = (document.getElementsByClassName('_19hf'))[0]; scrollingDiv.scrollIntoView(1)");
            sleep($$refCONFIG{'TIME_TO_WAIT'});
          } else { last; }
        }
      # Scroll chat
      } elsif ($typeHandle > 4) {
        my $maxScrollChatByDate = $$refWinConfig->rbMaxScrollChatByDate->Checked();
        my $maxDate;
        my $maxScrollChat;
        if ($maxScrollChatByDate) {
          my ($d, $m, $y) = $$refWinConfig->dtMaxScrollChatByDate->GetDate();
          $maxDate        = timelocal(0,0,0,$d,$m-1,$y); # Store in Unixtime format
        } else { $maxScrollChat = $$refWinConfig->tfMaxScrollChat->Text(); }
        if ($typeHandle == 5) { # Scroll chat to the top
          while (1) {
            $count++;
            # Scroll again
            if ($maxScrollChatByDate and $maxDate) { # Stop by date
              my $firstDisplayedDate;
              my $firstDisplayedDateCode = ($mech->selector('time._3oh-'))[0];
              if ($firstDisplayedDateCode->{innerHTML} =~ /(\d{2}\/\d{2}\/\d{4} \d{1,2}\:\d{2}[ap]m)/) {
                my $dateStr = $1;
                my $strp = DateTime::Format::Strptime->new(pattern => '%m/%d/%Y %I:%M%p');
                my $dt   = $strp->parse_datetime($dateStr);
                $firstDisplayedDate = timelocal(0,0,0,$dt->day(),$dt->month()-1,$dt->year());
              }
              last if $firstDisplayedDate <= $maxDate;
            } elsif ($maxScrollChat and $count > $maxScrollChat) { last; } # Stop by page
            if (($mech->selector('div._2k8v', any => 1))[0]) {
              $mech->eval_in_page("var scrollingDiv = (document.getElementsByClassName('_2k8v'))[0]; scrollingDiv.scrollIntoView(1)");
            } else { last; }
          }
        } elsif ($typeHandle == 6) { # Load Newer Msg
          while (1) {
            $count++;
            # Scroll again
            if ($maxScrollChatByDate and $maxDate) { # Stop by date
              my $lastDisplayedDate = ($mech->selector('h4._497p._2lpt time._3oh-'))[-1];
              if ($lastDisplayedDate->{innerHTML} =~ /(\d{2}\/\d{2}\/\d{4} \d{1,2}\:\d{2}[ap]m)/) {
                my $strp     = DateTime::Format::Strptime->new(pattern => '%m/%d/%Y %l:%M%p');
                my $dt       = $strp->parse_datetime($1);
                my $dateOnly = timelocal(0,0,0,$dt->day(),$dt->month()-1,$dt->year());
                last if $dateOnly >= $maxDate;
              }
            } elsif ($maxScrollChat and $count > $maxScrollChat) { last; } # Stop by page
            my $loadNewerCode = ($mech->selector('button._3quh._30yy._2t_._41jf'))[1];
            if ($loadNewerCode) {
              $loadNewerCode->click();
              sleep($$refCONFIG{'TIME_TO_WAIT'});
            # If No Load Older button, Load Newer if the first button
            } elsif ($loadNewerCode = ($mech->selector('button._3quh._30yy._2t_._41jf'))[0]) {
              $loadNewerCode->click();
              sleep($$refCONFIG{'TIME_TO_WAIT'});
            } else { last; }
          }
        } elsif ($typeHandle == 7) { # Load Older Msg
          while (1) {
            $count++;
            # Scroll again
            if ($maxScrollChatByDate and $maxDate) { # Stop by date
              my $firstDisplayedDate = ($mech->selector('h4._497p._2lpt time._3oh-'))[0];
              if ($firstDisplayedDate->{innerHTML} =~ /(\d{2}\/\d{2}\/\d{4} \d{1,2}\:\d{2}[ap]m)/) {
                my $strp     = DateTime::Format::Strptime->new(pattern => '%m/%d/%Y %l:%M%p');
                my $dt       = $strp->parse_datetime($1);
                my $dateOnly = timelocal(0,0,0,$dt->day(),$dt->month()-1,$dt->year());
                last if $dateOnly <= $maxDate;
              }
            } elsif ($maxScrollChat and $count > $maxScrollChat) { last; } # Stop by page
            my @loadOlderCode = $mech->selector('button._3quh._30yy._2t_._41jf');
            if (scalar(@loadOlderCode) > 1) {
              $loadOlderCode[0]->click();
              sleep($$refCONFIG{'TIME_TO_WAIT'});
            } else { last; }
          }
        # Remove
        } elsif ($typeHandle >= 8 and $typeHandle <= 12) {
          if ($typeHandle == 8 or $typeHandle == 12) {  # Remove Header
            $mech->eval_in_page('var div = document.getElementById("pagelet_bluebar"); if (div) { div.parentNode.removeChild(div); }'); # Blue bar
            $mech->eval_in_page('var div = document.getElementById("timeline_top_section"); if (div) { div.parentNode.removeChild(div); }'); # People Profile header
          }
          if ($typeHandle == 9 or $typeHandle == 12) { # Remove Left column
            $mech->eval_in_page('var div = document.getElementById("entity_sidebar"); if (div) { div.parentNode.removeChild(div); }'); # Event
            $mech->eval_in_page('var div = document.getElementById("leftCol"); if (div) { div.parentNode.removeChild(div); }'); # Group and page
            $mech->eval_in_page('var div = document.getElementById("u_0_16"); if (div) { div.parentNode.removeChild(div); }'); # People profile
          }
          if ($typeHandle == 10 or $typeHandle == 12) { # Remove Right menu
            $mech->eval_in_page('var div = document.getElementById("u_0_n"); if (div) { div.parentNode.removeChild(div); }'); # Event
            $mech->eval_in_page('var div = document.getElementById("u_0_s"); if (div) { div.parentNode.removeChild(div); }'); # Page
            $mech->eval_in_page('var div = document.getElementById("rightCol"); if (div) { div.parentNode.removeChild(div); }'); # Group
            $mech->eval_in_page('var div = document.getElementsByClassName("_14i5"); if (div[0]) { div[0].style.right = "0px"; div[0].style.left = "0px"; }'); # Modify scrollable area css
          }
          if ($typeHandle == 11 or $typeHandle == 12) { # Remove Bottom
            $mech->eval_in_page('var div = document.getElementById("pagelet_sidebar"); if (div) { div.parentNode.removeChild(div); }'); # All profile types
            $mech->eval_in_page('var div = document.getElementById("pagelet_dock"); if (div) { div.parentNode.removeChild(div); }'); # All profile types
          }
        }
      }
    $$refWin->Tray->Change(-tip => "$processName[$typeHandle] $$refSTR{'finished'}");
    if (!$$refWin->IsVisible()) {
      $$refWin->Tray->Change(-balloon_icon => 'info', -balloon_title => 'ExtractFace',
                             -balloon_tip => "$processName[$typeHandle] $$refSTR{'finished'}");
      $$refWin->Tray->ShowBalloon(1);
    }
    $$refWin->ChangeCursor($$refARROW);
  } else { Win32::GUI::MessageBox($$refWinPb2, $$refSTR{'warn4'}, $$refSTR{'Error'}, 0x40010); }
  # Turn off progress bar
  $$refWinPb2->lblPbCurr->Text('');
  $$refWinPb2->lblCount->Text('');
  &winPb2_Terminate;

}  #--- End handlePageThr

#--------------------------#
sub getListChats
#--------------------------#
{
  # Local variables
  my ($refTHR, $refARROW, $refHOURGLASS, $refWinChat, $refWinConfig, $refCONFIG, $CONFIG_FILE,
      $PROGDIR, $USERDIR, $DEBUG_FILE, $refWin, $refSTR) = @_;
  # Deal with crash
  $SIG{__DIE__} = sub {
    my $msgErr = $_[0];
    chomp($msgErr);
    $msgErr =~ s/[\t\r\n]/ /g;
    if ($msgErr =~ /NS_ERROR_FILE_IS_LOCKED/) { # Restart a new thread to continue
      $$refTHR = threads->create(\&getListChats, $refTHR, $refARROW, $refHOURGLASS, $refWinChat, $refWinConfig,
                                 $refCONFIG, $CONFIG_FILE, $PROGDIR, $USERDIR, $DEBUG_FILE, $refWin, $refSTR);
    } else {
      &debug($msgErr, $DEBUG_FILE) if $$refCONFIG{'DEBUG_LOGGING'};
      $$refWinChat->lblInProgress->Text('');
      $$refWin->ChangeCursor($$refARROW);
      $$refWin->Tray->Change(-tip => 'ExtractFace');
			my $err = (split(/ at /, $msgErr))[0];
      Win32::GUI::MessageBox($$refWinChat, "$$refSTR{'processCrash'}: $err", $$refSTR{'Error'}, 0x40010);
    }
  };
  $$refWin->ChangeCursor($$refHOURGLASS);
  my $mech;
  eval { $mech = WWW::Mechanize::Firefox->new(tab => 'current'); };
  if ($@) {
    if ($@ =~ /Failed to connect to/) {
      $$refWinChat->btnChatDumpNow->Disable();
      $$refWinChat->btnChatAddQueue->Disable();
      Win32::GUI::MessageBox($$refWinChat, $$refSTR{'errMozRepl'}, $$refSTR{'Error'}, 0x40010);
    }
    threads->exit();
  }
  if ($mech->uri() =~ /facebook.com/) {
    my $currAccountCode = $mech->selector('a._2s25', any => 1);
    if ($currAccountCode->{href} =~ /\/([^\/]+)$/) {
      my $title = $1;
      $title =~ s/[\#\<\>\:\"\/\\\|\?\*]/_/g;
      if ($title =~ /id=(\d+)/) { $title = $1; }
      $$refWinChat->tfChatName->Text("$title - $$refSTR{'Chat'}");
      $$refWinChat->tfChatCurrURL->Text($mech->uri);
    } else {
      $$refWinChat->btnChatDumpNow->Disable();
      $$refWinChat->btnChatAddQueue->Disable();
      Win32::GUI::MessageBox($$refWinChat, $$refSTR{'warn3'}, $$refSTR{'Warning'}, 0x40010);
      $$refWinChat->lblInProgress->Text('');
      $$refWin->ChangeCursor($$refARROW);
      threads->exit();
    }
  } else {
    $$refWinChat->btnChatDumpNow->Disable();
    $$refWinChat->btnChatAddQueue->Disable();
    Win32::GUI::MessageBox($$refWinChat, $$refSTR{'warn4'}, $$refSTR{'Error'}, 0x40010);
    $$refWinChat->lblInProgress->Text('');
    $$refWin->ChangeCursor($$refARROW);
    threads->exit();
  }
  # Scroll the Contacts list (if option is selected)
  if ($$refCONFIG{'AUTO_LOAD_SCROLL'}) {
    $$refWinChat->lblInProgress->Text("$$refSTR{'Scrolling'} $$refSTR{'Contacts'}...");
    while (1) {
      if (my $loadMoreContacts = ($mech->selector('div._19hf a', any => 1))[0]) {
        $mech->eval_in_page("var scrollingDiv = (document.getElementsByClassName('_19hf'))[0]; scrollingDiv.scrollIntoView(1)");
        sleep($$refCONFIG{'TIME_TO_WAIT'});
      } else { last; }
    }
  }
  # Save and parse the page
  $$refWinChat->lblInProgress->Text($$refSTR{'Parsing'}.'...');
  my $tempDir = "$USERDIR\\temp";
  mkdir("$tempDir") if !-d "$tempDir";
  my $htmlPage = "$tempDir\\temp.html";
  my $status = $mech->save_content($htmlPage);
  while ($status->{currentState} != $status->{PERSIST_STATE_FINISHED}) { sleep(1); }
  # Load content from file
  if (-T $htmlPage and open(my $fhTemp, "<:encoding(utf8)", $htmlPage)) {
    my $file_as_string = do { local $/ = <$fhTemp> };
    $file_as_string =~ s/[\r\n]//g;
    close($fhTemp);
    my @contactsNodes = split(/row_header_id_user/, $file_as_string);
    shift(@contactsNodes);
    # Parse each contact
    foreach my $contactNode (@contactsNodes) {
      if ($contactNode =~ /gridcell/ and $contactNode =~ /^:([^\"]+)\"/) {
        my $id = $1;
        my $name;
        my $url;
        if ($contactNode =~ /<span[^\>]+class="[^\"]*_1ht6[^\"]*"[^\>]*>([^\<]+)/) { $name = $1; }
        if ($contactNode =~ /data-href="([^\"]+)"/) {
          $url = $1;
          $url =~ s/&amp;/&/g;
        }
        if ($name and $url) {
          # Add to grid
          my $encodedName = encode($$refCONFIG{'CHARSET'}, $name);
          my $filename = $$refWinChat->tfChatName->Text() . " - $id";
          if (my $i = $$refWinChat->GridChats->InsertRow($encodedName, -1)) {
            $$refWinChat->GridChats->SetCellText($i, 0, ''        );
            $$refWinChat->GridChats->SetCellType($i, 0, GVIT_CHECK);
            $$refWinChat->GridChats->SetCellCheck($i, 0, 1);
            $$refWinChat->GridChats->SetCellText($i, 1, $encodedName);
            $$refWinChat->GridChats->SetCellText($i, 2, $filename);
            $$refWinChat->GridChats->SetCellText($i, 3, $url);
            $$refWinChat->GridChats->Refresh();
          }
        }
      }
    }
    $$refWinChat->GridChats->AutoSize();
    $$refWinChat->GridChats->ExpandLastColumn();
    $$refWinChat->GridChats->BringWindowToTop();
    $$refWinChat->lblInProgress->Text('');
  }
  $$refWin->ChangeCursor($$refARROW);
  &isDumpChatReady($refWinChat);
  
}  #--- End getListChats

#--------------------------#
sub addDumpChatToQueue
#--------------------------#
{
  # Local variables
  my ($refDumpParams, $refWinQueue, $USERDIR) = @_;
  # Create database dump file and add to queue grid
  if (&createDumpDB("$USERDIR\\Queue\\$$refDumpParams{processName}-$$refDumpParams{procID}\.db", $refDumpParams)) {
    if (&addToQueue($refWinQueue, "$$refDumpParams{processName}-$$refDumpParams{procID}", $$refDumpParams{filename},
                    $$refDumpParams{startingURL})) {
      return(1);
    } else { return(0);  }
  }
  
}  #--- End addDumpChatToQueue

#--------------------------#
sub loadQueue
#--------------------------#
{
  # Local variables
  my ($refWinQueue, $USERDIR, $refWin, $refSTR) = @_;
  # Verify the queue directory
  my $isQueue = 0;
  if (-d "$USERDIR\\Queue\\" and opendir(DIR, "$USERDIR\\Queue\\")) {
    while (readdir(DIR)) { if (/\.db$/) { $isQueue++; last; } }
    close(DIR);
  }
  # There is at least one pending job in queue
  if ($isQueue) {
    my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'pendingJob'} ?\r\n\r\n$$refSTR{'pendingJobWarn'}.", $$refSTR{'Queue'}, 0x40024);
    &createWinQueue() if !$$refWinQueue and $answer == 6;
    opendir(DIR, "$USERDIR\\Queue\\");
    while (my $file = readdir(DIR)) {
      if ($file =~ /\.db$/) {
        if ($answer == 6) { # Answer is yes, load the jobs in queue
          my $dbFile        = "$USERDIR\\Queue\\$file";
          my $dsn           = "DBI:SQLite:dbname=$dbFile";
          if (-f $dbFile and my $dbh = DBI->connect($dsn, undef, undef, { sqlite_unicode => 1})) {
            my $process = (split/\./, $file)[0];
            my $name    = $dbh->selectrow_array('SELECT value FROM INFOS WHERE key = ?', undef, 'filename');
            my $url     = $dbh->selectrow_array('SELECT value FROM INFOS WHERE key = ?', undef, 'startingURL');
            &addToQueue($refWinQueue, $process, $name, $url);
            undef $dbh;
          }
        } else { unlink("$USERDIR\\Queue\\$file"); } # Answer is no, delete the jobs
      }
    }
    close(DIR);
  }
  
}  #--- End loadQueue

#--------------------------#
sub existsInQueue
#--------------------------#
{
  # Local variables
  my ($refWinQueue, $filename) = @_;
  # Check in Queue Grid
  if ($$refWinQueue->gridQueue->GetRows() > 1) {
    for (my $i = 1; $i < $$refWinQueue->gridQueue->GetRows(); $i++) {
      return(1) if $$refWinQueue->gridQueue->GetCellText($i, 1) eq $filename;
    }
  }
  return(0);
  
}  #--- End existsInQueue

#--------------------------#
sub addToQueue
#--------------------------#
{
  # Local variables
  my ($refWinQueue, $process, $name, $url) = @_;
  # Add to Grid
  if (my $newLine = $$refWinQueue->gridQueue->InsertRow($process, -1)) {
    $$refWinQueue->gridQueue->SetCellText($newLine, 0, $process );
    $$refWinQueue->gridQueue->SetCellText($newLine, 1, $name );
    $$refWinQueue->gridQueue->SetCellText($newLine, 2, $url);
    $$refWinQueue->gridQueue->Refresh();
    $$refWinQueue->gridQueue->AutoSize();
    $$refWinQueue->gridQueue->ExpandLastColumn();
    $$refWinQueue->btnQueueProcess->Enable();
    return(1);
  } else { return(0); }
  
}  #--- End addToQueue

#--------------------------#
sub createDumpDB
#--------------------------#
{
  # Local variables
  my ($dbFile, $refInfos) = @_;
  my $dsn = "DBI:SQLite:dbname=$dbFile";
  my $dbh = DBI->connect($dsn, undef, undef, { AutoCommit => 1, sqlite_unicode => 1}) or return(0);
  # Create Infos table
  my $stmt = qq(CREATE TABLE IF NOT EXISTS INFOS
                (key            VARCHAR(255)  NOT NULL,
                 value          INT,
                 PRIMARY KEY (key)));
  my $rv = $dbh->do($stmt);
  return(0) if $rv < 0;
  # Insert all parameters
  my $sthUptInfos = $dbh->prepare('INSERT OR REPLACE INTO INFOS (key, value) VALUES(?,?)');
  foreach my $key (keys %{$refInfos}) { $sthUptInfos->execute($key, $$refInfos{$key}); }
  $sthUptInfos->finish();
  undef $dbh;
  return(1);
  
}  #--- End createDumpDB

#--------------------------#
sub createAlbumsDB
#--------------------------#
{
  # Local variables
  my ($dbFile, $refAlbums) = @_;
  my $dsn = "DBI:SQLite:dbname=$dbFile";
  my $dbh = DBI->connect($dsn, undef, undef, { AutoCommit => 0, sqlite_unicode => 1}) or return(0);
  # Create Infos table
  my $stmt = qq(CREATE TABLE IF NOT EXISTS ALBUMS
                (id           VARCHAR(255)  NOT NULL,
                 name         VARCHAR(255)  NOT NULL,
                 url          VARCHAR(255)  NOT NULL,
                 tmpPage      VARCHAR(255),
                 path         VARCHAR(255),
                 step         INT,
                 PRIMARY KEY (id)));
  my $rv = $dbh->do($stmt);
  return(0) if $rv < 0;
  # Insert all parameters
  my $sthUptAlbums = $dbh->prepare('INSERT OR REPLACE INTO ALBUMS (id, name, url) VALUES(?,?,?)');
  foreach my $albumId (keys %{$refAlbums}) { $sthUptAlbums->execute($albumId, $$refAlbums{$albumId}{name}, $$refAlbums{$albumId}{url}); }
  $sthUptAlbums->finish();
  $dbh->commit();
  undef $dbh;
  return(1);
  
}  #--- End createAlbumsDB

#--------------------------#
sub loadPage
#--------------------------#
{
  # Local variables
  my ($refMech, $url, $timeToWait) = @_;
  $$refMech->get($url, synchronize => 0);
  sleep($timeToWait);
  my $currURL = $$refMech->uri();
  my $currTitle;
  if ($currURL and $currURL =~ /https:\/\/(?:www|web).facebook.com\//) {
    if    ($currURL =~ /https:\/\/(?:www|web).facebook.com\/profile.php\?id=([^\/\&]+)/) { $currTitle = $1; }
    elsif ($currURL =~ /https:\/\/(?:www|web).facebook.com\/([^\/\?]+)/                ) { $currTitle = $1; }
  }
  return($currURL, $currTitle);

}  #--- End loadPage

#--------------------------#
sub formatDate
#--------------------------#
{
  # Local variables
  my $unixtime = shift;
  # Convert to string, local timezone
  if ($unixtime =~ /\./) { $unixtime = (split(/\./, $unixtime))[0]; }
  my ($s,$min,$hr,$d,$m,$y,$weekday,$ha,$isDST) = localtime($unixtime);
	return(sprintf("%04d\-%02d\-%02d %02d:%02d:%02d", $y+1900, $m+1, $d, $hr, $min, $s));

}  #--- End formatDate

#--------------------------#
sub scrollPage
#--------------------------#
{
  # Local variables
  my ($refMech, $time) = @_;
  # Scrolling down and wait for content to load
  $$refMech->eval_in_page('window.scrollTo(0,document.body.scrollHeight)');
  sleep($time);
  # Evaluate end of the page
  my ($end, $type) = $$refMech->eval_in_page('(window.innerHeight + window.scrollY) >= document.body.offsetHeight');
  if ($end == 1) {
    sleep($time); # Wait another X seconds and evaluate again
    ($end, $type) = $$refMech->eval_in_page('(window.innerHeight + window.scrollY) >= document.body.offsetHeight');
    return(1) if $end == 1; # End of the page
  }

}  #--- End scrollPage

#--------------------------#
sub scrollToBottom
#--------------------------#
{
  # Local variables
  my ($refMech, $time, $count, $refWinConfig) = @_;
  my $maxScrollByDate = $$refWinConfig->rbMaxScrollByDate->Checked();
  my $maxDate;
  my $maxScroll;
  if ($maxScrollByDate) {
    my ($d, $m, $y)   = $$refWinConfig->dtMaxScrollByDate->GetDate();
    $maxDate          = timelocal(0,0,0,$d,$m-1,$y); # Store in Unixtime format
  } else { $maxScroll = $$refWinConfig->tfMaxScroll->Text(); }
  while (1) { # Scrolling down and wait for content to load
    $$refMech->eval_in_page('window.scrollTo(0,document.body.scrollHeight)');
    $count++;
    sleep($time);
    # Scroll again
    if ($maxScrollByDate and $maxDate) { # Stop by date
      my $lastDisplayedDate = ($$refMech->selector('a._5pcq abbr'))[-1];
      if ($lastDisplayedDate->{outerHTML} =~ /data-utime="([^\"]+)"/) {
        my $date = $1;
        return(1) if $date <= $maxDate;
      }
    } elsif ($maxScroll and $count >= $maxScroll) { return(1); } # Stop by page
    # Evaluate end of the page
    my ($end, $type) = $$refMech->eval_in_page('(window.innerHeight + window.scrollY) >= document.body.offsetHeight');
    if ($end) { # End of the page, done ? Really ? Wait a bit more
      sleep($time);
      $$refMech->eval_in_page('window.scrollTo(0,document.body.scrollHeight)');
      ($end, $type) = $$refMech->eval_in_page('(window.innerHeight + window.scrollY) >= document.body.offsetHeight');
      return(1) if $end;
    }
  }

}  #--- End scrollToBottom

#--------------------------#
sub selectCatFriendPage
#--------------------------#
{
  # Local variables
  my ($refMech, $cat, $charSet) = @_;
  my @links = $$refMech->selector('div._3dc.lfloat._ohe._5brz a');
  foreach my $link (@links) {
		my $currentName = $link->{name};
		my $catName = encode($charSet, $currentName);
		$catName =~ s/[\<\>\:\"\/\\\|\?\*\.]/_/g;
    if ($catName eq $cat) {
      $link->click();
      return(1);
    }
  }

}  #--- End selectCatFriendPage

#--------------------------#
sub expandContent
#--------------------------#
{
  # Local variables
  my ($refMech, $refCONFIG) = @_;
  # Continue Reading
  $$refMech->eval_in_page("var el = document.getElementsByClassName('text_exposed_link'); for (var i=0;i<el.length; i++) { el[i].click(); }");
  # See more
  if ($$refCONFIG{'EXPAND_SEE_MORE'}) {
    $$refMech->eval_in_page("var el = document.getElementsByClassName('see_more_link'); for (var i=0;i<el.length; i++) { el[i].click(); }");
    $$refMech->eval_in_page("var el = document.getElementsByClassName('_5v47 fss'); for (var i=0;i<el.length; i++) { el[i].click(); }");
    $$refMech->eval_in_page("var el = document.getElementsByClassName('UFIReplySocialSentenceLinkText UFIReplySocialSentenceVerified'); for (var i=0;i<el.length; i++) { el[i].click(); }");
  }
  # More post (wait to find the good classname)
  # View \d+ more comments? / View previous comments / Reply / etc
  if ($$refCONFIG{'EXPAND_MORE_POSTS'}) {
    $$refMech->eval_in_page("var el = document.getElementsByClassName('UFIPagerLink'); for (var i=0;i<el.length; i++) { el[i].click(); }");
    $$refMech->eval_in_page("var el = document.getElementsByClassName('UFICommentLink'); for (var i=0;i<el.length; i++) { el[i].click(); }");
    $$refMech->eval_in_page("var el = document.getElementsByClassName('UFIBlingBox uiBlingBox feedbackBling'); for (var i=0;i<el.length; i++) { el[i].click(); }");
  }
  # See translation
  if ($$refCONFIG{'SEE_TRANSLATION'}) {
    # Translate comment
    $$refMech->eval_in_page("var el = document.getElementsByClassName('UFITranslateLink'); for (var i=0;i<el.length; i++) { el[i].click(); }");
    # Translate post
    my @links = $$refMech->selector('div._43f9 a');
    foreach (@links) { $_->click; }
  }

}  #--- End expandContent

#--------------------------#
sub exploreDir
#--------------------------#
{
	# Open Window Explorer
	my $dir = shift;
  Win32::Process::Create(my $ProcessObj, "$ENV{'WINDIR'}\\explorer.exe", "explorer $dir", 0, NORMAL_PRIORITY_CLASS, ".") if $dir and -d $dir;
	
}  #--- End exploreDir

#--------------------------#
sub debug
#--------------------------#
{
  # Local variables
  my ($refMsg, $DEBUG_FILE) = @_;
  my $dateStr = &formatDate(time);  
  # Save error msg in debug log file
  if (-e $DEBUG_FILE) { open(DEBUG,">>$DEBUG_FILE"); }
  else                { open(DEBUG,">$DEBUG_FILE");  }
  flock(DEBUG, 2);
  print DEBUG "$dateStr\t$refMsg\n";
  close(DEBUG);  

}  #--- End debug

#--------------------------#
sub rememberPosWin
#--------------------------#
{
  # Local variables
  my ($refSelWin, $selWin, $refWinConfig, $refCONFIG, $CONFIG_FILE) = @_;
	# Remember position
  my $winLeft = $$refSelWin->AbsLeft();
  my $winTop  = $$refSelWin->AbsTop();
  $$refCONFIG{$selWin.'_LEFT'} = $winLeft;
  $$refCONFIG{$selWin.'_TOP'}  = $winTop;
  &saveConfig($refCONFIG, $CONFIG_FILE);
  
}  #--- End rememberPosWin

#--------------------------#
sub saveConfig
#--------------------------#
{
  # Local variables
  my ($refConfig, $CONFIG_FILE) = @_;
  # Save configuration hash values
  open(CONFIG,">$CONFIG_FILE");
  flock(CONFIG, 2);
  foreach my $cle (keys %{$refConfig}) { print CONFIG "$cle = $$refConfig{$cle}\n"; }
  close(CONFIG);  

}  #--- End saveConfig

#--------------------------#
sub loadConfig
#--------------------------#
{
  # Local variables
  my ($refWinConfig, $refConfig, $CONFIG_FILE, $refWin) = @_;
  # If ini file exists
  if (-T $CONFIG_FILE) {
    # Open and load config values
    open(CONFIG, $CONFIG_FILE);
    my @tab = <CONFIG>;
    close(CONFIG);
    foreach (@tab) {
      chomp($_);
      my ($key, $value) = split(/ = /, $_);
      $$refConfig{$key}  = $value if $key;
    }
  }
  # Start minimized
  if (exists($$refConfig{'START_MINIMIZED'}))   { $$refWin->chStartMinimized->Checked($$refConfig{'START_MINIMIZED'});                  }
  else                                          { $$refWin->chStartMinimized->Checked(0);                                               } # Default is not checked
  # General settings
  if (exists($$refConfig{'AUTO_UPDATE'}))       { $$refWinConfig->chAutoUpdate->Checked($$refConfig{'AUTO_UPDATE'});                    }
  else                                          { $$refWinConfig->chAutoUpdate->Checked(1);   $$refConfig{'AUTO_UPDATE'} = 1;           } # Default is checked
  if (exists($$refConfig{'REMEMBER_POS'}))      { $$refWinConfig->chRememberPos->Checked($$refConfig{'REMEMBER_POS'});                  }
  else                                          { $$refWinConfig->chRememberPos->Checked(0);                                            } # Default is not checked
  if (exists($$refConfig{'DYNAMIC_MENU'}))      { $$refWinConfig->chOptDynamicMenu->Checked($$refConfig{'DYNAMIC_MENU'});               }
  else                                          { $$refWinConfig->chOptDynamicMenu->Checked(0); $$refConfig{'DYNAMIC_MENU'} = 0;        } # Default is not checked
  if (exists($$refConfig{'TIME_TO_WAIT'}))      { $$refWinConfig->upTimeToWait->SetPos($$refConfig{'TIME_TO_WAIT'});                    }
  else                                          { $$refWinConfig->upTimeToWait->SetPos(2); $$refConfig{'TIME_TO_WAIT'} =  2;            } # Default value is 2
  if (exists($$refConfig{'CHARSET'}))           { $$refWinConfig->cbCharset->SetCurSel($$refWinConfig->cbCharset->FindString($$refConfig{'CHARSET'})); }
  else                                          { $$refWinConfig->cbCharset->SetCurSel(0);    $$refConfig{'CHARSET'} = 'cp1252';        } # Default is cp1252
  if (exists($$refConfig{'DEBUG_LOGGING'}))     { $$refWinConfig->chDebugLogging->Checked($$refConfig{'DEBUG_LOGGING'});                }
  else                                          { $$refWinConfig->chDebugLogging->Checked(0); $$refConfig{'DEBUG_LOGGING'} = 0;         } # Default is not checked
  # Scroll and expand options
  if (exists($$refConfig{'MAX_LOADING_MSG'}))   { $$refWinConfig->upMaxScrollChat->SetPos($$refConfig{'MAX_LOADING_MSG'});              }
  else                                          { $$refWinConfig->upMaxScrollChat->SetPos(0); $$refConfig{'MAX_LOADING_MSG'} = 0;       } # Default value is 0 (No maximum)
  if (exists($$refConfig{'MAX_SCROLL'}))        { $$refWinConfig->upMaxScroll->SetPos($$refConfig{'MAX_SCROLL'});                       }
  else                                          { $$refWinConfig->upMaxScroll->SetPos(0); $$refConfig{'MAX_SCROLL'} = 0;                } # Default value is 0 (No maximum)
  if (exists($$refConfig{'EXPAND_SEE_MORE'}))   { $$refWinConfig->chOptSeemore->Checked($$refConfig{'EXPAND_SEE_MORE'});                }
  else                                          { $$refWinConfig->chOptSeemore->Checked(1);   $$refConfig{'EXPAND_SEE_MORE'} = 1;       } # Default is checked
  if (exists($$refConfig{'EXPAND_MORE_POSTS'})) { $$refWinConfig->chOptPosts->Checked($$refConfig{'EXPAND_MORE_POSTS'});                }
  else                                          { $$refWinConfig->chOptPosts->Checked(1);     $$refConfig{'EXPAND_MORE_POSTS'} = 1;     } # Default is checked
  if (exists($$refConfig{'SEE_TRANSLATION'}))   { $$refWinConfig->chOptTranslate->Checked($$refConfig{'SEE_TRANSLATION'});              }
  else                                          { $$refWinConfig->chOptTranslate->Checked(0); $$refConfig{'SEE_TRANSLATION'} = 0;       } # Default is not checked
  # Dump options
  if (exists($$refConfig{'AUTO_LOAD_SCROLL'}))  { $$refWinConfig->chOptAutoLoadScroll->Checked($$refConfig{'AUTO_LOAD_SCROLL'});        }
  else                                          { $$refWinConfig->chOptAutoLoadScroll->Checked(1); $$refConfig{'AUTO_LOAD_SCROLL'} = 1; } # Default is checked
  if (exists($$refConfig{'OPT_SCROLL_TOP'}))    { $$refWinConfig->chOptScrollTop->Checked($$refConfig{'OPT_SCROLL_TOP'});               }
  else                                          { $$refWinConfig->chOptScrollTop->Checked(1); $$refConfig{'OPT_SCROLL_TOP'} = 1;        } # Default is checked
  if (exists($$refConfig{'REMEMBER_SAVE_DIR'})) { $$refWinConfig->chRememberSaveDir->Checked($$refConfig{'REMEMBER_SAVE_DIR'});         }
  else                                          { $$refWinConfig->chRememberSaveDir->Checked(1); $$refConfig{'REMEMBER_SAVE_DIR'} = 1;  } # Default value is checked
  if (exists($$refConfig{'SILENT_PROGRESS'}))   { $$refWinConfig->chSilentProgress->Checked($$refConfig{'SILENT_PROGRESS'});            }
  else                                          { $$refWinConfig->chSilentProgress->Checked(1); $$refConfig{'SILENT_PROGRESS'} = 1;     } # Default value is checked
  if (exists($$refConfig{'OPEN_REPORT'}))       { $$refWinConfig->chOptOpenReport->Checked($$refConfig{'OPEN_REPORT'});                 }
  else                                          { $$refWinConfig->chOptOpenReport->Checked(1); $$refConfig{'OPEN_REPORT'} = 1;          } # Default value is checked
  if (exists($$refConfig{'DONT_OPEN_REPORT'}))  { $$refWinConfig->chOptDontOpenReport->Checked($$refConfig{'DONT_OPEN_REPORT'});        }
  else                                          { $$refWinConfig->chOptDontOpenReport->Checked(1); $$refConfig{'DONT_OPEN_REPORT'} = 1; } # Default value is checked
  if (exists($$refConfig{'CLOSE_USED_TABS'}))   { $$refWinConfig->chCloseUsedTabs->Checked($$refConfig{'CLOSE_USED_TABS'});             }
  else                                          { $$refWinConfig->chCloseUsedTabs->Checked(1); $$refConfig{'CLOSE_USED_TABS'} = 1;      } # Default is checked
  if (exists($$refConfig{'DEL_TEMP_FILES'}))    { $$refWinConfig->chDelTempFiles->Checked($$refConfig{'DEL_TEMP_FILES'});               }
  else                                          { $$refWinConfig->chDelTempFiles->Checked(1); $$refConfig{'DEL_TEMP_FILES'} = 1;        } # Default is checked
  &saveConfig($refConfig, $CONFIG_FILE);

}  #--- End loadConfig

#--------------------------#
sub checkUpdate
#--------------------------#
{
  # Local variables
  my ($confirm, $VERSION, $refWin, $refSTR) = @_;
  # Download the version file  
  my $ua = new LWP::UserAgent;
  $ua->agent("ExtractFaceUpdate $VERSION");
  $ua->default_header('Accept-Language' => 'en');
  my $req = new HTTP::Request GET => $URL_VER;
  my $res = $ua->request($req);
  # Success, compare versions
  if ($res->is_success) {
    my $status  = $res->code;
    my $content = $res->content;
    my $currVer;
    $currVer = $1 if $content =~ /([\d\.]+)/i;
    # No update available
    if ($currVer le $VERSION) {
      Win32::GUI::MessageBox($$refWin, $$refSTR{'update1'}, $$refSTR{'update2'}, 0x40040) if $confirm; # Up to date
    } else {
      $$refWin->Show();
      # Download with browser
      my $answer = Win32::GUI::MessageBox($$refWin, "$$refSTR{'update4'} $currVer $$refSTR{'update5'} ?",
                                          $$refSTR{'update3'}, 0x40024);
      if ($answer == 6) { # Open Firefox to XL-Tools page
        $$refWin->ShellExecute('open', $URL_TOOL,'','',1) or
				Win32::GUI::MessageBox($$refWin, Win32::FormatMessage(Win32::GetLastError()),
                               "$$refSTR{'update3'} ExtractFace", 0x40010);
      }
    }
  } else { # Error 
    my $status  = $res->code;
    my $error   = $res->status_line;
    Win32::GUI::MessageBox($$refWin, "$$refSTR{'Error'}: $$refSTR{'returnedCode'} = [$status]; $$refSTR{'returnedError'} = [$error].",
                           $$refSTR{'errConnection'}, 0x40010);
  }

}  #--- End checkUpdate

#------------------------------------------------------------------------------#
1;