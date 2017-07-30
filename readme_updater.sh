#!/bin/bash

GH_REPO="@github.com/moj-analytical-services/xltabr.git"
FULL_REPO="https://$GH_TOKEN$GH_REPO"

mkdir out
cd out

# Checkout repo so we have all the files we need to build
git init
git remote add origin $FULL_REPO
git fetch
git config user.name "xltabr-travis"
git config user.email "travis"
git checkout dev

# On the dev branch we want to convert vignettes/readme.Rmd to readme.md
Rscript -f travis.R

git add README.md
git commit -m "readme update by travis after $TRAVIS_COMMIT"
git push origin dev
