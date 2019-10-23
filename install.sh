apt-get install build-essential
apt-get install curl
curl -L http://cpanmin.us | perl - App::cpanminus
cpanm LWP::UserAgent
cpanm JSON::MaybeXS
cpanm Azure::AD::Auth
cpanm JSON

# Install Tk
apt-get install libx11-dev
apt-get install perl-tk
cpanm Tk::GridColumns
