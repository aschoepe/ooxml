#!/bin/sh

PACKAGE_NAME=$1
PACKAGE_VERSION=$2
PKG_TCL_SOURCES=$3

mkdir -p uv/${PACKAGE_NAME}${PACKAGE_VERSION}

cp ${PKG_TCL_SOURCES} pkgIndex.tcl uv/${PACKAGE_NAME}${PACKAGE_VERSION}
cd uv
zip -qr ${PACKAGE_NAME}${PACKAGE_VERSION}.zip ${PACKAGE_NAME}${PACKAGE_VERSION}
tar -czf ${PACKAGE_NAME}${PACKAGE_VERSION}.tar.gz ${PACKAGE_NAME}${PACKAGE_VERSION}
rm -fr ${PACKAGE_NAME}${PACKAGE_VERSION}

cp template-download.html download.html
ex - download.html <<!
g/PACKAGE_VERSION/s//${PACKAGE_VERSION}/g
wq
!

cd ../..
tar --exclude=uv/\* -czf ${PACKAGE_NAME}/uv/${PACKAGE_NAME}${PACKAGE_VERSION}-src.tar.gz ${PACKAGE_NAME}
