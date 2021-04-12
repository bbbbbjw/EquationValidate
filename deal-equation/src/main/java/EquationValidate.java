import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class EquationValidate {
    public EquationValidate() {
    }

    public void validateEquation(String path) throws Exception {
        System.out.println("------------开始检测--------------");
        int chapterCnt=0;
        int indexCnt=0;
        XWPFDocument word = new XWPFDocument(new FileInputStream(path));
        try {
            List<XWPFParagraph> paragraphs = word.getParagraphs();
            int iii=0;
            for (XWPFParagraph paragraph : paragraphs) {
                //System.out.println("iii:"+iii++);
                String pStyleVal="0";
                Node pgnode= paragraph.getCTP().getDomNode();
                Node ppr = getChildNode(pgnode, "w:pPr");
                if(ppr!=null){
                    Node pStyle = getChildNode(ppr, "w:pStyle");
                    if(pStyle!=null){
                        pStyleVal = pStyle.getAttributes().item(0).getNodeValue();
                        //System.out.println("pStyleVal"+pStyleVal);
                    }
                }
                if (pStyleVal.equals("1")){
                    chapterCnt++;
                    indexCnt=0;
                    //System.out.println("chapterCnt"+chapterCnt);
                    //System.out.println("indexCnt"+indexCnt);
                }
                StringBuffer text = new StringBuffer();
                List<XWPFRun> runs = paragraph.getRuns();
                boolean hasEquation=false;
                for (XWPFRun run : runs) {

                    Node runNode = run.getCTR().getDomNode();
                    text.append(getText(runNode));

                    String math = getMath(run, runNode);
                    if (math!="")hasEquation=true;
                    text.append(math);
                }
                if(hasEquation){
                    indexCnt++;
                    String indexStr="("+chapterCnt+"."+indexCnt+")";
                    //检查制表符设置
                    Node tabsNode = getChildNode(pgnode, "w:tabs");
                    NodeList tabsNodeChildNodes = tabsNode.getChildNodes();
                    Map<String,String> tabs=new HashMap<>();
                    for(int i=0;i<tabsNodeChildNodes.getLength();i++){
                        Node node=tabsNodeChildNodes.item(i);
                        String val=node.getAttributes().item(0).getNodeValue();
                        String pos = node.getAttributes().item(1).getNodeValue();
                        tabs.put(val,pos);
                    }
                    if(tabs.get("right")!=null&&tabs.get("center")!=null){
                        double dr=Double.parseDouble(tabs.get("right"));
                        double dc=Double.parseDouble(tabs.get("center"));
                        if(dr/dc>2.3||dr/dc<1.7) System.out.println("居中有问题");
                    }
                    else{
                        System.out.println("居中有问题");
                    }
                    //检查前后是否是tab
                    Node pre=null;
                    int v1=0;
                    for (XWPFRun run : runs) {
                        Node runNode = run.getCTR().getDomNode();
                        String math = getMath(run, runNode);
                        if(v1==1){
                            v1=0;
                            if(getChildNode(runNode,"w:tab")==null) System.out.println("error");
                        }
                        if (math!="") {
                            if(getChildNode(pre,"w:tab")==null) System.out.println("error");
                            v1=1;//通知后一个node进行检查
                        }
                        pre=runNode;
                    }
                    String s = text.toString();
                    String[] strings = s.split("mmm公式mmm");
                    if(!(strings[0].equals("  解:")||strings[0].equals("  假定")))
                        System.out.println("公式前不准有“解:”，“假定”以外的文字,您有文本为："+strings[0]);
                    if(!strings[1].equals(indexStr)) System.out.println("章节编号不对，文中为:"+strings[1]+"应为:"+indexStr);
                }

            }
        } finally {
            word.close();
            System.out.println("检测完成");
        }
    }

    /**
     * 获取字符串
     *
     * @param runNode
     * @return
     */
    private String getText(Node runNode) {
        Node textNode = getChildNode(runNode, "w:t");
        if (textNode == null) {
            return "";
        }
        return textNode.getFirstChild().getNodeValue();
    }

    private String getMath(XWPFRun run, Node runNode) throws Exception {
        Node objectNode = getChildNode(runNode, "w:object");
        if (objectNode == null) {
            return "";
        }
        Node shapeNode = getChildNode(objectNode, "v:shape");
        if (shapeNode == null) {
            return "";
        }
        Node imageNode = getChildNode(shapeNode, "v:imagedata");
        if (imageNode == null) {
            return "";
        }
        Node binNode = getChildNode(objectNode, "o:OLEObject");
        if (binNode == null) {
            return "";
        }

        XWPFDocument word = run.getDocument();
        return "mmm公式mmm";
    }

    private Node getChildNode(Node node, String nodeName) {
        if (!node.hasChildNodes()) {
            return null;
        }
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (nodeName.equals(childNode.getNodeName())) {
                return childNode;
            }
            childNode = getChildNode(childNode, nodeName);
            if (childNode != null) {
                return childNode;
            }
        }
        return null;
    }

}
