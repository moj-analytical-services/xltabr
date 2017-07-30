#!/bin/bash

GH_REPO="git@github.com:moj-analytical-services/xltabr.git"

# Start ssh agent and add key
eval "$(ssh-agent -s)" # Start the ssh agent
chmod 600 deploy_rsa
ssh-add deploy_rsa

mkdir out
cd out


# Checkout repo so we have all the files we need to build

git config user.name "xltabr-travis"
git config user.email "travis"

git clone $GH_REPO
cd xltabr
git checkout dev

# On the dev branch we want to convert vignettes/readme.Rmd to readme.md
Rscript travis.R

git add README.md
git commit -m "readme update by travis after $TRAVIS_COMMIT"
git push origin dev
