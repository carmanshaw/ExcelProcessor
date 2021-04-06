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
	
	//源表格控件srclist和数据srcdlm
	private DefaultListModel srcdlm = new DefaultListModel();
	final JList<String> srclist = new JList<String>();
	private ExcelApi srcExcelApi = new ExcelApi();
	
	//目的表格控件distlist和数据distdlm
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

        // 窗体大小
        this.setSize(900, 350);
        // 屏幕显示初始位置
        this.setLocation(200, 200);
       
        // 显示
        this.setVisible(true);
        // 退出窗体后将JFrame同时关闭
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }
    
    public void initListEvent() {
    	// 添加选项选中状态被改变的监听器
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
     * *创建中间功能按钮区域布局
     * @return
     */
    private JPanel createFuncLayout() {
    	JButton jbEmerge,jbCompare;
    	
    	JPanel panel = new JPanel(new GridLayout(2,1));
    	
    	jbEmerge = new JButton("<html> >>>> <br>(合并)</html>");
    	jbEmerge.setPreferredSize(new Dimension(100, 100));
    	jbCompare = new JButton("比较");
    	jbCompare.setPreferredSize(new Dimension(100, 100));
    	
    	jbEmerge.addActionListener(new ActionListener() {
					
					@Override
					public void actionPerformed(ActionEvent e) {
						// TODO Auto-generated method stub
						
						if(srcExcelApi.getSelectedIndex() == -1) {
							showHintDialog("左侧未选择内容!");
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
							
							//重置目的文件相关内容
							distExcelApi = null;
							distExcelApi = new ExcelApi();
							updateListCont(distExcelApi,path,jtxtfDistFileName,distdlm,distlist);
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						showHintDialog("合并完成!");
					}
				});
		
    	jbCompare.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				if(srcExcelApi.getSelectedIndex() == -1) {
					showHintDialog("左侧未选择内容!");
					return;
				}
				
				if(distExcelApi.getSelectedIndex() == -1) {
					showHintDialog("右侧未选择内容!");
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
		    	
		    	showHintDialog("数据匹配完成,全部相同!");
			}
		});

    	panel.add(jbEmerge,BorderLayout.NORTH);
    	panel.add(jbCompare,BorderLayout.SOUTH);
    	return panel;
    	
    }
    
    /**
     * *通用列表创建 
     * @param lst
     * @param df
     * @param excelApi
     * @return
     */
    private JPanel createListLayout(JList lst) {
    	JPanel panel = new JPanel();
        // 设置一下首选大小
        lst.setPreferredSize(new Dimension(300, 200));

        panel.add(lst);
        return panel;
    	
    }
    
    /**
     *  *创建的文本路径输入框布局
     * @return
     */
    private JPanel createFileInputHeader() {
    	JPanel jpSrcFileInput, jpDistFileInput,jpFileLayout;
        JButton jbSrcFileInput, jbDistFileInput;
        JLabel jlbSrcFileName,jlbDistFileName;
    	JTextField jtxtfSrcFileName;//,jtxtfDistFileName;
    	// JPanel布局默认是FlowLayout流布局
    	jpSrcFileInput = new JPanel();
    	jpDistFileInput = new JPanel();
    	jpFileLayout = new JPanel();//jpFileLayout = new JPanel(new GridLayout(2,1));

    	jlbSrcFileName = new JLabel("源文件路径:");
    	jtxtfSrcFileName = new JTextField(20);
    	jtxtfSrcFileName.setEditable(false);
    	jbSrcFileInput = new JButton("选择文件");
    	jlbDistFileName = new JLabel("目标文件路径:");
//    	jtxtfDistFileName = new JTextField(20);
    	jtxtfDistFileName.setEditable(false);
    	jbDistFileInput = new JButton("选择文件");
    	
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

        // 设置布局管理器(Jpanel默认流布局)
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
     * *加载表格显示的内容
     * @param jtxtfile
     * @param dlm
     * @param jlist
     */
    private void updateListCont(ExcelApi excelApi,String path,JTextField jtxtfile,DefaultListModel dlm,JList jlist) {
    	
		if(path != null) {
			jtxtfile.setText(path);
			boolean bret = excelApi.loadFile(path);
			if(!bret) {
				showHintDialog("表格打开失败!");
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
     * *文件选择框输入
     *  true  源文件  false 目的文件
     * @return
     */
    private File fileInputDialog() {
    	JFileChooser jfc = new JFileChooser();// 文件选择器
    	jfc.setCurrentDirectory(new File("E:\\Project\\ExcelPro\\ExcelProcessor\\ExcelProcessorTool"));// 文件选择器的初始目录定为d盘
//    	jfc.setCurrentDirectory(new File("D:\\"));// 文件选择器的初始目录定为d盘
    	jfc.setFileSelectionMode(0);// 设定只能选择到文件
		int state = jfc.showOpenDialog(null);// 此句是打开文件选择器界面的触发语句
		if (state == 1) {
			return null;// 撤销则返回
		} else {
			return jfc.getSelectedFile();// f为选择到的文件
//			return f.getAbsolutePath();
		}
    }
    
    /**
     * *弹框提示
     * @param cont
     */
    public void showHintDialog(String cont) {
		  // 创建一个新窗口
		  JFrame newJFrame = new JFrame("提示");
		 
		  newJFrame.setSize(300, 100);
		 
		  // 把新窗口的位置设置到 relativeWindow 窗口的中心
		  newJFrame.setLocationRelativeTo(this);
		 
		  // 点击窗口关闭按钮, 执行销毁窗口操作（如果设置为 EXIT_ON_CLOSE, 则点击新窗口关闭按钮后, 整个进程将结束）
		  newJFrame.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
		 
		  // 窗口设置为不可改变大小
		  newJFrame.setResizable(false);
		 
		  JPanel panel = new JPanel(new GridLayout(1, 1));
		 
		  // 在新窗口中显示一个标签
		  JLabel label = new JLabel(cont);
//		  label.setFont(new Font(null, Font.PLAIN, 15));
		  label.setHorizontalAlignment(SwingConstants.CENTER);
		  label.setVerticalAlignment(SwingConstants.CENTER);
		  panel.add(label);
		 
		  newJFrame.setContentPane(panel);
		  newJFrame.setVisible(true);
	 }
}
