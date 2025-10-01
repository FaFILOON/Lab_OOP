package p10_1;

import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.BorderLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.Popup;

public class Calculator extends JFrame implements ActionListener {

    private final JButton[] buttons = {
        new JButton("%"), new JButton("CE"), new JButton("C"), new JButton("/"),
        new JButton("7"), new JButton("8"), new JButton("9"), new JButton("*"),
        new JButton("4"), new JButton("5"), new JButton("6"), new JButton("-"),
        new JButton("1"), new JButton("2"), new JButton("3"), new JButton("+"),
        new JButton("."), new JButton("0"), new JButton("<-"), new JButton("=")
    };
    private final JTextField field;
    private double num1 = 0, num2 = 0, result = 0;
    private char operator = '\0';

    public Calculator(String title) {
        super(title);
        setSize(300, 400);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        field = new JTextField();
        field.setFont(new Font("Consolas", Font.PLAIN, 40));
        field.setEditable(false);
        add(field, BorderLayout.NORTH);

        JPanel buttonPanel = new JPanel();
        buttonPanel.setLayout(new GridLayout(5, 4));

        for (int i = 0; i < 20; i++) {
            buttonPanel.add(buttons[i]);
            buttons[i].addActionListener(this);
        }
        add(buttonPanel, BorderLayout.CENTER);
        
        setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        String command = e.getActionCommand();

        if ((command.charAt(0) >= '0' && command.charAt(0) <= '9') || command.equals(".")) {
        	if (field.getText().equals("0") && !command.equals(".")) {
                field.setText(command);
            } else {
                field.setText(field.getText() + command);
            }
        } else if (command.equals("C")) {
            field.setText("");
            num1 = 0;
            num2 = 0;
            result = 0;
        } else if (command.equals("<-")) {
        	String temp = field.getText();
        	if (temp.length() > 0) {
                field.setText(temp.substring(0, temp.length() - 1));
            }
        } else if (command.equals("CE")) {
                field.setText("0");
        } else if (command.equals("=")) {
        	if (!field.getText().isEmpty() && operator != '\0') {
                num2 = Double.parseDouble(field.getText());
                switch (operator) {
                    case '+': result = num1 + num2; break;
                    case '-': result = num1 - num2; break;
                    case '*': result = num1 * num2; break;
                    case '/': result = num1 / num2; break;
                    case '%': result = num1 % num2; break;
                }
                field.setText(String.valueOf(result));
                num1 = result;
                operator = '\0';
            }
        } else {
            if (operator != '\0') {
                buttons[19].doClick();
            }
            num1 = Double.parseDouble(field.getText());
            operator = command.charAt(0);
            field.setText("0");
        }
    }
}
