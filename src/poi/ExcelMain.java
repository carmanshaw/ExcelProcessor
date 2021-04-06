package poi;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.List;

import javax.swing.Box;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.ListModel;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.WindowConstants;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Font;

public class ExcelMain extends JFrame {
	
	//Դ���ؼ�srclist������srcdlm
	private DefaultListModel srcdlm = new DefaultListModel();
	final JList<String> srclist = new JList<String>();
	private ExcelApi srcExcelApi = new ExcelApi();
	
	//Ŀ�ı��ؼ�distlist������distdlm
	private DefaultListModel distdlm = new DefaultListModel();
	final JList<String> distlist = new JList<String>();
	private ExcelApi distExcelApi = new ExcelApi();
	JTextField jtxtfDistFileName = new JTextField(20);;

    public static void main(String[] args) {
        new ExcelMain();
    }

    public ExcelMain() {
    	    	
    	this.add(createFileInputHeader(),BorderLayout.NORTH);
    	
    	JPanel panel = new JPanel();
    	panel.add(createListLayout(srclist),BorderLayout.WEST);
    	panel.add(createFuncLayout(),BorderLayout.CENTER);
    	panel.add(createListLayout(distlist),BorderLayout.EAST);
    	
    	initListEvent();
    	
    	this.add(panel,BorderLayout.CENTER);
    	
    	JPanel versPanel = new JPanel(new BorderLayout());  
    	versPanel.add(new JLabel("<html>Vers: V1.00.2104061,Authored By Mr.Boxin</html>",JLabel.CENTER),BorderLayout.CENTER); 
    	this.add(versPanel,BorderLayout.SOUTH);

        // �����С
        this.setSize(900, 350);
        // ��Ļ��ʾ��ʼλ��
        this.setLocation(200, 200);
       
        // ��ʾ
        this.setVisible(true);
        // �˳������JFrameͬʱ�ر�
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }
    
    public void initListEvent() {
    	// ���ѡ��ѡ��״̬���ı�ļ�����
        srclist.addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                int index = srclist.getSelectedIndex();
                srcExcelApi.setSelectedIndex(index);
            }
        });
        
        distlist.addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                int index = distlist.getSelectedIndex();
                distExcelApi.setSelectedIndex(index);
            }
        });
    }
    
    /**
     * *�����м书�ܰ�ť���򲼾�
     * @return
     */
    private JPanel createFuncLayout() {
    	JButton jbEmerge,jbCompare;
    	
    	JPanel panel = new JPanel(new GridLayout(2,1));
    	
    	jbEmerge = new JButton("<html> >>>> <br>(�ϲ�)</html>");
    	jbEmerge.setPreferredSize(new Dimension(100, 100));
    	jbCompare = new JButton("�Ƚ�");
    	jbCompare.setPreferredSize(new Dimension(100, 100));
    	
    	jbEmerge.addActionListener(new ActionListener() {
					
					@Override
					public void actionPerformed(ActionEvent e) {
						// TODO Auto-generated method stub
						
						if(srcExcelApi.getSelectedIndex() == -1) {
							showHintDialog("���δѡ������!");
							return;
						}

						int srcRows = srcExcelApi.getRows();
				    	for(int i=0;i<srcRows;i++) {
				    		Row srcrow = srcExcelApi.getRowLine(i);
				    		String srcmainkey = (String) srcExcelApi.getCellFormatValue(srcrow.getCell(0));
				    		if(srcmainkey != null) {
				    			distExcelApi.appendCont(srcmainkey, srcrow.getCell(srcExcelApi.getSelectedIndex()));
				    		}
				    	}
				    	
				    	String[] split = distExcelApi.getFilePath().split("\\.");
						String path = split[0] + "_Merge." + split[1];
						
						try {
							distExcelApi.save(path);
							
							//����Ŀ���ļ��������
							distExcelApi = null;
							distExcelApi = new ExcelApi();
							updateListCont(distExcelApi,path,jtxtfDistFileName,distdlm,distlist);
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						showHintDialog("�ϲ����!");
					}
				});
		
    	jbCompare.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if(srcExcelApi.getSelectedIndex() == -1) {
					showHintDialog("���δѡ������!");
					return;
				}
				
				if(distExcelApi.getSelectedIndex() == -1) {
					showHintDialog("�Ҳ�δѡ������!");
					return;
				}
				
				int srcRows = srcExcelApi.getRows();
		    	for(int i=1;i<srcRows;i++) {
		    		Row srcrow = srcExcelApi.getRowLine(i);
		    		String srcmainkey = (String) srcExcelApi.getCellFormatValue(srcrow.getCell(0));
		    		if(srcmainkey != null) {
		    			Cell srcCell = srcrow.getCell(srcExcelApi.getSelectedIndex());
		    			boolean bret = distExcelApi.compare(srcmainkey, srcCell);
		    			if(!bret) {
//		    				srcExcelApi.setFontStyle(srcCell, Font.COLOR_RED);
		    				showHintDialog(distExcelApi.getCompareError());
		    				return ;
		    			}
		    		}
		    	}
		    	
		    	showHintDialog("����ƥ�����,ȫ����ͬ!");
			}
		});

    	panel.add(jbEmerge,BorderLayout.NORTH);
    	panel.add(jbCompare,BorderLayout.SOUTH);
    	return panel;
    	
    }
    
    /**
     * *ͨ���б��� 
     * @param lst
     * @param df
     * @param excelApi
     * @return
     */
    private JPanel createListLayout(JList lst) {
    	JPanel panel = new JPanel();
        // ����һ����ѡ��С
        lst.setPreferredSize(new Dimension(300, 200));

        panel.add(lst);
        return panel;
    	
    }
    
    /**
     *  *�������ı�·������򲼾�
     * @return
     */
    private JPanel createFileInputHeader() {
    	JPanel jpSrcFileInput, jpDistFileInput,jpFileLayout;
        JButton jbSrcFileInput, jbDistFileInput;
        JLabel jlbSrcFileName,jlbDistFileName;
    	JTextField jtxtfSrcFileName;//,jtxtfDistFileName;
    	// JPanel����Ĭ����FlowLayout������
    	jpSrcFileInput = new JPanel();
    	jpDistFileInput = new JPanel();
    	jpFileLayout = new JPanel();//jpFileLayout = new JPanel(new GridLayout(2,1));

    	jlbSrcFileName = new JLabel("Դ�ļ�·��:");
    	jtxtfSrcFileName = new JTextField(20);
    	jtxtfSrcFileName.setEditable(false);
    	jbSrcFileInput = new JButton("ѡ���ļ�");
    	jlbDistFileName = new JLabel("Ŀ���ļ�·��:");
//    	jtxtfDistFileName = new JTextField(20);
    	jtxtfDistFileName.setEditable(false);
    	jbDistFileInput = new JButton("ѡ���ļ�");
    	
    	jbSrcFileInput.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				File file = fileInputDialog();
				if(file != null) {
					updateListCont(srcExcelApi,file.getAbsolutePath(),jtxtfSrcFileName,srcdlm,srclist);
				}
			}
		});
    	jbDistFileInput.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				File file = fileInputDialog();
				if(file != null) {
					updateListCont(distExcelApi,file.getAbsolutePath(),jtxtfDistFileName,distdlm,distlist);
				}
			}
		});

        // ���ò��ֹ�����(JpanelĬ��������)
    	jpSrcFileInput.add(jlbSrcFileName);
    	jpSrcFileInput.add(jtxtfSrcFileName);
    	jpSrcFileInput.add(jbSrcFileInput);
    	
    	jpDistFileInput.add(jlbDistFileName);
    	jpDistFileInput.add(jtxtfDistFileName);
    	jpDistFileInput.add(jbDistFileInput);

    	jpFileLayout.add(jpSrcFileInput, BorderLayout.NORTH);
    	jpFileLayout.add(jpDistFileInput, BorderLayout.CENTER);
    	return jpFileLayout;
    }
    
    /**
     * *���ر����ʾ������
     * @param jtxtfile
     * @param dlm
     * @param jlist
     */
    private void updateListCont(ExcelApi excelApi,String path,JTextField jtxtfile,DefaultListModel dlm,JList jlist) {
    	
		if(path != null) {
			jtxtfile.setText(path);
			boolean bret = excelApi.loadFile(path);
			if(!bret) {
				showHintDialog("����ʧ��!");
				return;
			}
			dlm.clear();
			List<String> titles = excelApi.getTitles();
			for(int i=0;i<titles.size();i++) {
				dlm.add(i, titles.get(i));
			}
			jlist.setModel(dlm);
		}
    }
    
    /**
     * *�ļ�ѡ�������
     *  true  Դ�ļ�  false Ŀ���ļ�
     * @return
     */
    private File fileInputDialog() {
    	JFileChooser jfc = new JFileChooser();// �ļ�ѡ����
    	jfc.setCurrentDirectory(new File("E:\\Project\\ExcelPro\\ExcelProcessor\\ExcelProcessorTool"));// �ļ�ѡ�����ĳ�ʼĿ¼��Ϊd��
//    	jfc.setCurrentDirectory(new File("D:\\"));// �ļ�ѡ�����ĳ�ʼĿ¼��Ϊd��
    	jfc.setFileSelectionMode(0);// �趨ֻ��ѡ���ļ�
		int state = jfc.showOpenDialog(null);// �˾��Ǵ��ļ�ѡ��������Ĵ������
		if (state == 1) {
			return null;// �����򷵻�
		} else {
			return jfc.getSelectedFile();// fΪѡ�񵽵��ļ�
//			return f.getAbsolutePath();
		}
    }
    
    /**
     * *������ʾ
     * @param cont
     */
    public void showHintDialog(String cont) {
		  // ����һ���´���
		  JFrame newJFrame = new JFrame("��ʾ");
		 
		  newJFrame.setSize(300, 100);
		 
		  // ���´��ڵ�λ�����õ� relativeWindow ���ڵ�����
		  newJFrame.setLocationRelativeTo(this);
		 
		  // ������ڹرհ�ť, ִ�����ٴ��ڲ������������Ϊ EXIT_ON_CLOSE, �����´��ڹرհ�ť��, �������̽�������
		  newJFrame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		 
		  // ��������Ϊ���ɸı��С
		  newJFrame.setResizable(false);
		 
		  JPanel panel = new JPanel(new GridLayout(1, 1));
		 
		  // ���´�������ʾһ����ǩ
		  JLabel label = new JLabel(cont);
//		  label.setFont(new Font(null, Font.PLAIN, 15));
		  label.setHorizontalAlignment(SwingConstants.CENTER);
		  label.setVerticalAlignment(SwingConstants.CENTER);
		  panel.add(label);
		 
		  newJFrame.setContentPane(panel);
		  newJFrame.setVisible(true);
	 }
}
