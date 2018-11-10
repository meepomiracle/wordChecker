package com.xysy;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CommentsDocument;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class XWPFExtendDocument extends XWPFDocument {

    private MyXWPFCommentsDocument myXWPFCommentsDocument;

    private BigInteger commentId = BigInteger.ZERO;

    public BigInteger getCommentId() {
        return commentId;
    }

    public void setCommentId(BigInteger commentId) {
        this.commentId = commentId;
    }

    public XWPFExtendDocument(InputStream is) throws InvalidFormatException, IOException {
        super(is);
        this.myXWPFCommentsDocument = createCommentsDocument(this);
    }

    public MyXWPFCommentsDocument getMyXWPFCommentsDocument() {
        return myXWPFCommentsDocument;
    }

    public void setMyXWPFCommentsDocument(MyXWPFCommentsDocument myXWPFCommentsDocument) {
        this.myXWPFCommentsDocument = myXWPFCommentsDocument;
    }

    //第一个核心方法获取自己创建得自定义得对象
    private static MyXWPFCommentsDocument createCommentsDocument(XWPFDocument document) throws InvalidFormatException {
        OPCPackage oPCPackage = document.getPackage();
        PackagePartName partName = PackagingURIHelper.createPartName("/word/comments.xml");
        PackagePart part = oPCPackage.createPart(partName, "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml");
        MyXWPFCommentsDocument myXWPFCommentsDocument = new MyXWPFCommentsDocument(part);

        String rId = "rId" + (document.getRelationParts().size()+1);
        document.addRelation(rId, XWPFRelation.COMMENT, myXWPFCommentsDocument);

        return myXWPFCommentsDocument;
    }

    //第二个就是自己定义得核心对象
    public static class MyXWPFCommentsDocument extends POIXMLDocumentPart {

        private CTComments comments;

        private MyXWPFCommentsDocument(PackagePart part){
            super(part);
            comments = CommentsDocument.Factory.newInstance().addNewComments();
        }

        public CTComments getComments() {
            return comments;
        }

        @Override
        protected void commit() throws IOException {
            XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
            xmlOptions.setSaveSyntheticDocumentElement(new QName(CTComments.type.getName().getNamespaceURI(), "comments"));
            PackagePart part = getPackagePart();
            OutputStream out = part.getOutputStream();
            comments.save(out, xmlOptions);
            out.close();
        }

    }
}
