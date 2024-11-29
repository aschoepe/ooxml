#!/bin/sh

PACKAGE_NAME=$1
PACKAGE_VERSION=$2
PKG_TCL_SOURCES=$3

# ignore AppleDouble files
# macOS 10.4
COPY_EXTENDED_ATTRIBUTES_DISABLE=1
export COPY_EXTENDED_ATTRIBUTES_DISABLE
# macOS >= 10.5
COPYFILE_DISABLE=1
export COPYFILE_DISABLE

rm -f uv/ooxml*.zip uv/ooxml*.tar.gz

mkdir -p uv/${PACKAGE_NAME}${PACKAGE_VERSION}

cp ${PKG_TCL_SOURCES} manifest.txt pkgIndex.tcl uv/${PACKAGE_NAME}${PACKAGE_VERSION}
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
tar --exclude=uv/\* --exclude=doc/ECMA\* --exclude=ooxml/.fslckout --exclude=ooxml/.fossil-settings --exclude=ooxml/.gitattributes --exclude=ooxml/.vscode -czf ${PACKAGE_NAME}/uv/${PACKAGE_NAME}${PACKAGE_VERSION}-src.tar.gz ${PACKAGE_NAME}
