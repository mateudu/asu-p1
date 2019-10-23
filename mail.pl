#!/usr/local/bin/perl -w 

use Tk; 
use Tk::widgets qw/Dialog/; 
use Tk::GridColumns;
use subs qw/build_menubar fini/; 
use vars qw/$MW $VERSION/; 
use Azure::AD::Auth;
use Azure::AD::DeviceLogin;
use JSON;
use feature 'say';
use strict; 

my $Auth = Azure::AD::DeviceLogin->new(
        resource_id => 'https://graph.microsoft.com/',
        message_handler => sub { say $_[0] },
        client_id => '8d899e7a-9101-4dad-8104-34dd5348ae83',
        tenant_id => 'c944b87d-367e-4691-a6c3-3b14725a81dd',
    );

#login_to_azure_ad();

$MW = MainWindow->new;
$MW->title("Outlook Email Browser");
$MW->geometry( "=1280x768+100+100" );

my $CurrentFolder;

my $GridColumns;
my $Canvas;
my $EmailListData;

configure_main_menubar($MW);

$Canvas = $MW->Canvas;
$Canvas->pack(-expand => 1, -fill => 'both');

render_table();

MainLoop;

sub configure_main_menubar { 
    my ($mainWindow) = @_;
    
    my $menubar = $mainWindow->Menu; 
    $mainWindow->configure(-menu => $menubar);

    my $file = $menubar->cascade(-label => 'File'); 
    $file->command(-label => "New", -command => sub { open_email_write('', '', '') });
    $file->command(-label => "Exit", -command => sub { exit() });
    
    my $folders = $menubar->cascade(-label => 'Folders');
    for my $folder( @{get_json_folders()} ) {    
        $folders->command(-label => $folder->{displayName}, -command => sub { set_current_folder($folder) } );
    }

    $menubar;
}

sub fini { exit; }

sub login_to_azure_ad {
    say $Auth->access_token;
    $Auth->access_token;
}

sub get_json_folders {
    use HTTP::Tiny;
  my $ua = HTTP::Tiny->new;
  my $response = $ua->get(
    'https://graph.microsoft.com/v1.0/me/mailFolders', 
    {
       headers => { Authorization => 'Bearer ' . $Auth->access_token }
    }
  );
    decode_json($response->{content})->{'value'};
}

sub get_json_messages {
    my ( $folderId ) = @_;
    
    use HTTP::Tiny;
    my $ua = HTTP::Tiny->new;
    my $response = $ua->get(
    'https://graph.microsoft.com/v1.0/me/mailFolders/' . $folderId . '/messages?$select=sender,subject,receivedDateTime&$top=50', 
    {
       headers => { Authorization => 'Bearer ' . $Auth->access_token }
       #headers => { Authorization => 'Bearer ' . 'eyJ0eXAiOiJKV1QiLCJub25jZSI6InFUWDhFaHI0bXFSck9FbGc1dmxOa0pkd2hFM0JKeWU0aWpfZHU3Ukt0bDgiLCJhbGciOiJSUzI1NiIsIng1dCI6ImFQY3R3X29kdlJPb0VOZzNWb09sSWgydGlFcyIsImtpZCI6ImFQY3R3X29kdlJPb0VOZzNWb09sSWgydGlFcyJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMDg3NzBjMTUtMTJlMy00NjYyLWE2MTQtNWZkMDNkNzk2MmUwLyIsImlhdCI6MTU3MTUxNzIyMCwibmJmIjoxNTcxNTE3MjIwLCJleHAiOjE1NzE1MjExMjAsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVVFBdS84TkFBQUFLWWhLbkZoREdHTnhqdXFVUXZCVWR4aXJPUjAxemtaNmxDSTVoSSthZ2JFTyt1aURZeWxxenRVZjJaMzRSNEdzRjJ5QTc3eXFHTWxSQUhnUFdzTzBqQT09IiwiYW1yIjpbInB3ZCIsIm1mYSJdLCJhcHBfZGlzcGxheW5hbWUiOiJBU1UtTzM2NS1BcHAiLCJhcHBpZCI6IjJiODIyZDI1LTJiMzItNDhhMS05OGVhLTY5YThhZGQ5ZDJjNSIsImFwcGlkYWNyIjoiMCIsImZhbWlseV9uYW1lIjoiRHVkZWsiLCJnaXZlbl9uYW1lIjoiTWF0ZXVzeiIsImlwYWRkciI6IjE4OC4xNDYuMTg2LjUzIiwibmFtZSI6Ik1hdGV1c3ogRHVkZWsiLCJvaWQiOiJkOTFlMmY0MC05MjZmLTQ4ZGItYjY3MS05ZWE1ZTIzZDE1ZjAiLCJwbGF0ZiI6IjE0IiwicHVpZCI6IjEwMDM3RkZFQTM1RjFEQTIiLCJzY3AiOiJlbWFpbCBNYWlsLlJlYWQgTWFpbC5SZWFkV3JpdGUgTWFpbC5TZW5kIG9mZmxpbmVfYWNjZXNzIG9wZW5pZCBwcm9maWxlIFVzZXIuUmVhZCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6ImhHM0VOUWZfZDd0NmpOOE5DWDd0UmtRSUpqSHAzeVFvZnZRZ2NkSzh2aWMiLCJ0aWQiOiIwODc3MGMxNS0xMmUzLTQ2NjItYTYxNC01ZmQwM2Q3OTYyZTAiLCJ1bmlxdWVfbmFtZSI6Im1hdGV1c3ouZHVkZWtAY2xvdWRjb29raW5nLnBsIiwidXBuIjoibWF0ZXVzei5kdWRla0BjbG91ZGNvb2tpbmcucGwiLCJ1dGkiOiJ3VmprT19KVWYwdW9NX1pvMzd0MkFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiXSwieG1zX3RjZHQiOjE1MDA5MTMxOTN9.BUJueyczGSJ14omF7Xerq7h6JEUCmf2Ffq2673xGd481KH0gyVKvpqp5dL500tg3sTL1uzRXAT3ZUSqeQlIokLZhVajezRphW3UqYUyyUud1J6__vJwjUtY0pRvdp0zE1CCW1iKUmPGUdSrcRK2Mb6Y5vC7YRWltwUbOO2LhAgAnVkUy7DF_y25ub_vwv5y2Lx6cBh7pjjjKOWx6rOKhRhbVK_uLxaPhEnyRHfpv3-5m-D_9ylTDV2YwlJ9r4eZZXwNEBWQdDtq8ueiHIC13xqFUgUPVf4FZBTJ3ZlJOCVkD-OVdsYvX2YP4XvtLvjtRikHYVcgYvkD0vExc3UCzBw' }
    }
  );
    decode_json($response->{content})->{'value'};
}

sub get_json_message {
    my ( $messageId ) = @_;

    use HTTP::Tiny;
    my $ua = HTTP::Tiny->new;
    my $response = $ua->get(
    'https://graph.microsoft.com/v1.0/me/messages/' . $messageId . '?$select=id,body,subject,sender,toRecipients,ccRecipients,bccRecipients', 
    {
       headers => { Authorization => 'Bearer ' . $Auth->access_token }       
    }
  );
    decode_json($response->{content});
}

sub send_json_message {
    my ( $to, $subject, $body ) = @_;

    my $json = JSON->new;
my $data_to_json = {
    message=>{
            subject => $subject,
            body => {
                contentType => 'Text',
                content => $body
            },
            toRecipients => [
                {
                    emailAddress => {
                        address => $to
                    }
                }
            ]
    }
};

    use HTTP::Tiny;
    my $ua = HTTP::Tiny->new;
    my $response = $ua->post(
    'https://graph.microsoft.com/v1.0/me/sendMail', 
    {
       headers => { 
           Authorization => 'Bearer ' . $Auth->access_token,
           'Content-Type' => 'application/json'
       },
       content => $json->encode($data_to_json)
    }
    );

    return $response;
}

sub delete_json_message {
    my ( $emailId ) = @_;

    use HTTP::Tiny;
    my $ua = HTTP::Tiny->new;
    my $response = $ua->delete(
    'https://graph.microsoft.com/v1.0/me/messages/' . $emailId, 
    {
       headers => { 
           Authorization => 'Bearer ' . $Auth->access_token
       }
    }
    );

    return $response;
}

sub set_current_folder {
    my ( $folder ) = @_;
    
    $CurrentFolder = $folder;
    $EmailListData = get_json_messages($CurrentFolder->{id});     
    render_table();
}

sub render_table {
    if (defined $Canvas) {
        $Canvas->destroy();
    }

    $Canvas = $MW->Canvas;
    $Canvas->pack(-expand => 1, -fill => 'both');

    my $data = [];
    if (defined $CurrentFolder) {
        $data = [ map { [ $_->{sender}->{emailAddress}->{name}, $_->{subject}, $_->{receivedDateTime}, $_->{id} ] } @{ $EmailListData } ];        
    }
    
    $GridColumns = $Canvas->Scrolled(
        'GridColumns' =>
        -scrollbars => 'ose',
        -data => $data,
        -select_cmd => \&select_table_row,
        -columns => \my @columns        
    )->pack(
        -fill => 'both',
        -expand => 1        
    )->Subwidget( 'scrolled' ); # do not forget this one ;)
 
    @columns = (
        {
            -text => 'Sender',
            -command => $GridColumns->sort_cmd( 0, 'abc' ),
        },
        {
            -text => 'Subject',
            -command => $GridColumns->sort_cmd( 0, 'abc' ),
            -weight => 1,
        },
        {
            -text => 'Received at',
            -command => $GridColumns->sort_cmd( 1, 'abc' ),        
        },
    );
    
    $GridColumns->refresh;
}

sub select_table_row {
    my ( $a1, $a2, $row, $col ) = @_;
    
    if ($col == 0) {        
        open_email_preview($EmailListData->[$row]->{id});    
    }
}

sub convert_outlook_person_to_text {
    my ( $person ) = @_;

    return $person->{emailAddress}->{name} . ' (' . $person->{emailAddress}->{address} . ')';
}

sub get_multiple_people_text {
    my ( $people ) = @_;
    
    my $toTextValue = '';
    for my $person ( @{$people} ) {    
        if ($toTextValue ne '') {
            $toTextValue = $toTextValue . ', ';
        }

        $toTextValue = $toTextValue . convert_outlook_person_to_text($person);
    }

    return $toTextValue;
}

sub open_email_preview {
    my ( $emailId ) = @_;
    
    my $emailData = get_json_message($emailId);

    my $readEmailWindow = MainWindow->new;
    $readEmailWindow->title($emailData->{subject});
    $readEmailWindow->geometry( "=1280x768+100+100" );
    
    my $menubar = $readEmailWindow->Menu;     
    $menubar->command(-label=>"Delete", -command=> sub { handle_email_delete($readEmailWindow, $emailId) } );
    $menubar->command(-label=>"Reply", -command=> sub { open_email_write($emailData->{sender}->{emailAddress}->{address}, $emailData->{subject}, $emailData->{body}->{content}) } );
    $readEmailWindow->configure(-menu => $menubar);

    my $readEmailWindowCanvas = $readEmailWindow->Canvas;
    $readEmailWindowCanvas->pack(-expand => 1, -fill => 'both');

    my $fromFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_from = $fromFrame->Label(-text => 'FROM:')->pack (-side => 'left');        
    my $entry_from = $fromFrame->Label(-text => convert_outlook_person_to_text($emailData->{sender}))->pack (-side => 'left' );    

    my $toFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_to = $toFrame->Label(-text => 'TO:')->pack (-side => 'left');        
    my $entry_to = $toFrame->Label(-text => get_multiple_people_text($emailData->{toRecipients}))->pack (-side => 'left' );

    my $ccFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_cc = $ccFrame->Label(-text => 'CC:')->pack (-side => 'left');        
    my $entry_cc = $ccFrame->Label(-text => get_multiple_people_text($emailData->{ccRecipients}))->pack (-side => 'left' );

    my $bccFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_bcc = $bccFrame->Label(-text => 'BCC:')->pack (-side => 'left');        
    my $entry_bcc = $bccFrame->Label(-text => get_multiple_people_text($emailData->{bccRecipients}))->pack (-side => 'left' );
        
    my $content  = $readEmailWindowCanvas->Scrolled('Text');
    $content->pack(
        -fill => 'both',
        -expand => 1);
    $content->insert( 'end', $emailData->{body}->{content} );
}

sub open_email_write {
    my ( $toParameter, $subjectParameter, $contentParameter ) = @_;
    
    if ($subjectParameter ne '') {
        $subjectParameter = 'RE: ' . $subjectParameter;
    }

    if ($contentParameter ne '') {
        $contentParameter = "\n==========\n" . $contentParameter;
    }

    my $readEmailWindow = MainWindow->new;
    $readEmailWindow->title('New email');
    $readEmailWindow->geometry( "=1280x768+100+100" );

    my $readEmailWindowCanvas = $readEmailWindow->Canvas;
    $readEmailWindowCanvas->pack(-expand => 1, -fill => 'both');

    my $toFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_to = $toFrame->Label(-text => 'TO:')->pack (-side => 'left');        
    my $entry_to = $toFrame->Entry(-text => $toParameter);
    $entry_to->pack (-side => 'left', -fill => 'both', -expand => 1 );    

    my $subjectFrame = $readEmailWindowCanvas->Frame()->pack(-fill => 'both');
    my $label_subject = $subjectFrame->Label(-text => 'SUBJECT:')->pack (-side => 'left');        
    my $entry_subject = $subjectFrame->Entry(-text => $subjectParameter);
    $entry_subject->pack (-side => 'left', -fill => 'both', -expand => 1 );
        
    my $content = $readEmailWindowCanvas->Text();
    $content->insert('1.0', $contentParameter);
    $content->pack(
        -fill => 'both',
        -expand => 1);

    my $sendButton = $readEmailWindowCanvas->Button(-text => 'Send', -command => sub { 
        handle_email_send($readEmailWindow, $entry_to->get(), $entry_subject->get(), $content->get('1.0', 'end-1c')) 
        }
    )->pack(-fill => 'both');
}

sub handle_email_send {
    my ( $window, $to, $subject, $body ) = @_;
    
    my $response = send_json_message( $to, $subject, $body );

    if ($response->{success}) {
        my $popupResponse = $window->messageBox(
            -title => 'Email sent',
            -message => 'Email sent',
            -type => 'Ok'
        );
        $window->destroy();
    }
    else {
        my $popupResponse = $window->messageBox(
            -title => 'Email not sent',
            -message => 'Email not sent',
            -type => 'Ok'
        );        
    }    
}

sub handle_email_delete {
    my ( $window, $emailId ) = @_;
    
    my $response = delete_json_message( $emailId );

    if ($response->{success}) {
        my $popupResponse = $window->messageBox(
            -title => 'Email deleted',
            -message => 'Email deleted',
            -type => 'Ok'
        );
        $window->destroy();
    }
    else {
        my $popupResponse = $window->messageBox(
            -title => 'Email not deleted',
            -message => 'Email not deleted',
            -type => 'Ok'
        );        
    }    
}