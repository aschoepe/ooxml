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

# xattr -r -d com.apple.provenance .: recursively (-r) deletes the extended attribute com.apple.provenance from the specified file or directory. com.apple.provenance: An extended attribute used by macOS (e.g., Finder or Gatekeeper) to track the origin of a file, such as downloaded sources.
# tar --no-mac-metadata: Do not store or restore Mac extended metadata (e.g., resource forks, Finder info, and so forth). This metadata is stored in special “AppleDouble” files or as extended attributes. This option is useful when creating archives to be used on non-Mac platforms.
# tar --disable-copyfile: This option is equivalent to setting the environment variable COPYFILE_DISABLE=1. It disables copying of extended attributes and resource forks by preventing the use of the copyfile() API. Use this when you want to avoid creation of ._* AppleDouble files or when those are not needed.
# tar --no-xattrs: Do not store or restore extended file attributes (xattrs). These are often system or application metadata not relevant or supported on other systems.

cd uv
zip -x '.DS_Store' -x '*/.DS_Store' -qr ${PACKAGE_NAME}${PACKAGE_VERSION}.zip ${PACKAGE_NAME}${PACKAGE_VERSION}
tar --no-xattrs --no-mac-metadata --disable-copyfile --exclude='.DS_Store' --exclude='*/.DS_Store' -czf ${PACKAGE_NAME}${PACKAGE_VERSION}.tar.gz ${PACKAGE_NAME}${PACKAGE_VERSION}
rm -fr ${PACKAGE_NAME}${PACKAGE_VERSION}

cp template-download.html download.html
ex - download.html <<!
g/PACKAGE_VERSION/s//${PACKAGE_VERSION}/g
wq
!

cd ../..
tar --no-xattrs --no-mac-metadata --disable-copyfile --exclude='uv/*' --exclude='doc/ECMA*' --exclude='ooxml/.fslckout' --exclude='ooxml/.fossil-settings' --exclude='ooxml/.gitattributes' --exclude='ooxml/.vscode' --exclude='.DS_Store' --exclude='*/.DS_Store' --exclude='autom4te.cache' --exclude='configure~' --exclude='config.*' -czf ${PACKAGE_NAME}/uv/${PACKAGE_NAME}${PACKAGE_VERSION}-src.tar.gz ${PACKAGE_NAME}
