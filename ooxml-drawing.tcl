#
#  ooxml ECMA-376 Office Open XML File Formats
#  https://www.ecma-international.org/publications/standards/Ecma-376.htm
#
#  Copyright (C) 2025 Rolf Ade, DE
#  All rights reserved.
#
#  Redistribution and use in source and binary forms, with or without modification,
#  are permitted provided that the following conditions are met:
#
#  1. Redistributions of source code must retain the above copyright notice, this
#     list of conditions and the following disclaimer.
#  2. Redistributions in binary form must reproduce the above copyright notice,
#     this list of conditions and the following disclaimer in the documentation
#     and/or other materials provided with the distribution.
#  3. Neither the name of the project nor the names of its contributors may be used
#     to endorse or promote products derived from this software without specific
#     prior written permission.
#
#  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
#  EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
#  OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT
#  SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
#  INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
#  TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
#  BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
#  CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
#  ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF
#  SUCH DAMAGE.
#

package require Tcl 8.6.7-
package require tdom 0.9.6-
package require ooxml

namespace eval ::ooxml::drawing {

    variable xmlns

    array set xmlns {
        mc http://schemas.openxmlformats.org/markup-compatibility/2006
        o urn:schemas-microsoft-com:office:office
        r http://schemas.openxmlformats.org/officeDocument/2006/relationships
        rel http://schemas.openxmlformats.org/package/2006/relationships
        v urn:schemas-microsoft-com:vml
        w http://schemas.openxmlformats.org/wordprocessingml/2006/main
        w10 urn:schemas-microsoft-com:office:word
        w14 http://schemas.microsoft.com/office/word/2010/wordml
        wp http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing
        wp14 http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing
        wpg http://schemas.microsoft.com/office/word/2010/wordprocessingGroup
        wps http://schemas.microsoft.com/office/word/2010/wordprocessingShape

        
    }


    proc picture {} {
        variable properties
        
        upvar opts opts
        
    }
}
