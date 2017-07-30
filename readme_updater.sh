#!/bin/bash

GH_REPO="git@github.com:moj-analytical-services/xltabr.git"


mkdir out
cd out

# Start ssh agent and add key
eval "$(ssh-agent -s)" # Start the ssh agent
chmod 600 deploy_rsa
ssh-add deploy_rsa

# Checkout repo so we have all the files we need to build
git clone $GH_REPO
git config user.name "xltabr-travis"
git config user.email "travis"
git checkout dev

# On the dev branch we want to convert vignettes/readme.Rmd to readme.md
Rscript travis.R

git add README.md
git commit -m "readme update by travis after $TRAVIS_COMMIT"
git push origin dev
