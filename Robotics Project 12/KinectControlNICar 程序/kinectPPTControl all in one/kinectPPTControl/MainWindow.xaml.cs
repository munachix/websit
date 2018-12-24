using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Kinect;//引用Kinect的.NET文件
using System.Windows.Forms;
using kinectPPTControl;
using System.Runtime.InteropServices;

namespace kinectPPTControl
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public bool isBackGestureActive = false;
        public bool isForwardGestureActive = false;
        public bool isLeftGestureActive = false;
        public bool isRightGestureActive = false;//Four	Boolean	variables	are defined for the current record is what kind of attitude
        KinectSensor kinectSensor;
        private byte[] pixelData;//Define an array for storing the color image data 
        private Skeleton[] skeletonData;//Define an array for storing data bones 
        public double height;//Defined variable height（double type） 
        CenterPosition centerPosition;//definition CenterPositionVariablecenterPosition
        private readonly int MOUSEEVENTF_LEFTDOWN = 0x0002;//Simulation of the left mouse button pressed
        private readonly int MOUSEEVENTF_MOVE = 0x0001;//Simulate mouse movement 
        private readonly int MOUSEEVENTF_LEFTUP = 0x0004;//Simulation of the left mouse button lift

        [DllImport("user32")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);//Define a mouse event

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            kinectSensor = (from sensor in KinectSensor.KinectSensors where sensor.Status == KinectStatus.Connected select sensor).FirstOrDefault();//By selecting Kinect sensor to obtain the label first or default, then get Kinect sensor 
            kinectSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);//Itallows the sensor to acquire color image information, and attribute setting 
            kinectSensor.Start();//KinectPower
            kinectSensor.ColorFrameReady += kinectSensor_ColorFrameReady;//Capturing a color image
            kinectSensor.SkeletonStream.Enable();//Get information allows the sensor to the bone
            kinectSensor.SkeletonFrameReady += new EventHandler<SkeletonFrameReadyEventArgs>(kinectSensor_SkeletonFrameReady);//Get skeletal information
        }
        private void kinectSensor_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame imageFrame = e.OpenColorImageFrame())
            {
                if (imageFrame != null)
                {
                    this.pixelData = new byte[imageFrame.PixelDataLength];
                    imageFrame.CopyPixelDataTo(this.pixelData);
                    this.ColorImage.Source = BitmapSource.Create(imageFrame.Width, imageFrame.Height, 96, 96, PixelFormats.Bgr32, null, pixelData, imageFrame.Width * imageFrame.BytesPerPixel);//当彩色图像信息准备完成时，将其显示在Image1中
                }
            }
        }
        private void kinectSensor_SkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    skeletonData = new Skeleton[kinectSensor.SkeletonStream.FrameSkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(this.skeletonData);
                    Skeleton skeleton = (from s in skeletonData where s.TrackingState == SkeletonTrackingState.Tracked select s).FirstOrDefault();
                    if (skeleton != null)
                    {
                        SkeletonCanvas.Visibility = Visibility.Visible;
                        ProcessGesture(skeleton);
                        textBox1.Text = " Get Skeleton Data!! ";//When ready skeletal information, call the posture detection function ProcessGesture, and displays the information ready to complete skeleton of the prompt text in textBox1

                    }
                }
            }
        }

        private void ProcessGesture(Skeleton s)
        {
            Joint rightHand = (from j in s.Joints where j.JointType == JointType.HandRight select j).FirstOrDefault();
            Joint head = (from j in s.Joints where j.JointType == JointType.Head select j).FirstOrDefault();
            Joint centerShoulder = (from j in s.Joints where j.JointType == JointType.ShoulderCenter select j).FirstOrDefault();
            Joint rightfoot = (from j in s.Joints where j.JointType == JointType.FootRight select j).FirstOrDefault();//Being right hand, head, shoulder and foot bones midpoint of the original information
            RealPosition realPositionRightHand;
            RealPosition realPositionHead;
            RealPosition realPositionRightFoot;
            RealPosition realPositionCenterShoulder;//The definition of the right hand, head,	shoulder and foot bones midpoint of the true position variable (realPosition category)
            realPositionRightHand = new RealPosition(rightHand);
            realPositionHead = new RealPosition(head);
            realPositionRightFoot = new RealPosition(rightfoot);
            realPositionCenterShoulder = new RealPosition(centerShoulder);//The right hand, head, shoulder and foot bones midpoint of the original information into skeletal real location, and given to the appropriate properties (see specific process Class1.cs)
            height = realPositionHead.realY - realPositionRightFoot.realY;//Users measuring height
            centerPosition = new CenterPosition();
            centerPosition.X = realPositionCenterShoulder.realX + 0.1 * height;
            centerPosition.Y = realPositionCenterShoulder.realY - 0.1 * height;//Set reference point X, Y coordinate values
            textBox2.Text = Convert.ToString(height);
            if (realPositionRightHand.realX < centerPosition.X - height * 0.2)//Detecting whether the left hand
            {
                if (isLeftGestureActive)
                { }//If you have left is not issued a directive (here it is mainly to prevent repeat sending the same instruction, wasting CPU memory)
                else
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);//Click
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");//Delete an existing command
                    System.Windows.Forms.SendKeys.SendWait("3");//Enter the appropriate command
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");//Implementation of existing directives (repeated here mainly to prevent leakage of the key, because the situation of several key leak appeared in the previous debugging)
                    isForwardGestureActive = false;
                    isBackGestureActive = false;
                    isLeftGestureActive = true;
                    isRightGestureActive = false;//Update four Boolean
                }
            }
            if (realPositionRightHand.realX > centerPosition.X + height * 0.2)//检测右手是否向右
            {
                if (isRightGestureActive)//如果已经向右，则不做动作
                { }
                else
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);//单击
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");//将现有指令删除
                    System.Windows.Forms.SendKeys.SendWait("4");//输入相应的指令
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");//执行现有指令（这里的多次重复主要是为了防止漏键，因为此前的调试中出现过几次漏键的情况）
                    isForwardGestureActive = false;
                    isBackGestureActive = false;
                    isLeftGestureActive = false;
                    isRightGestureActive = true;//更新四个布尔量 
                }
            }
            if (realPositionRightHand.realX > centerPosition.X - height * 0.2 && realPositionRightHand.realX < centerPosition.X + height * 0.2 && realPositionRightHand.realY < centerPosition.Y + height * 0.2 && realPositionRightHand.realY > centerPosition.Y - height * 0.2)
            {
                if (!isForwardGestureActive && !isBackGestureActive &&!isLeftGestureActive &&!isRightGestureActive)
                { }
                else
                {
                    
                    
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);//单击
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");//将现有指令删除
                    System.Windows.Forms.SendKeys.SendWait("0");//输入相应的指令
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");//执行现有指令（这里的多次重复主要是为了防止漏键，因为此前的调试中出现过几次漏键的情况）
                    isForwardGestureActive = false;
                    isBackGestureActive = false;
                    isLeftGestureActive = false;
                    isRightGestureActive = false;//更新四个布尔量 
                }
            }

            if (realPositionRightHand.realY > centerPosition.Y + height * 0.2)
            {
                if (isForwardGestureActive)
                { }
                else
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);//单击
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                    System.Windows.Forms.SendKeys.SendWait("{Backspace}");//将现有指令删除
                    System.Windows.Forms.SendKeys.SendWait("1");//输入相应的指令
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");
                    System.Windows.Forms.SendKeys.SendWait("{Enter}");//执行现有指令（这里的多次重复主要是为了防止漏键，因为此前的调试中出现过几次漏键的情况）
                    isForwardGestureActive = true;
                    isBackGestureActive = false;
                    isLeftGestureActive = false;
                    isRightGestureActive = false;//更新四个布尔量 
                    
                }
            }
             if (realPositionRightHand.realY < centerPosition.Y - height * 0.2)
             {
                 if (isBackGestureActive)
                 { }
                 else
                 {
                     mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                     mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);//单击
                     System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                     System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                     System.Windows.Forms.SendKeys.SendWait("{Backspace}");
                     System.Windows.Forms.SendKeys.SendWait("{Backspace}");//将现有指令删除
                     System.Windows.Forms.SendKeys.SendWait("2");//输入相应的指令
                     System.Windows.Forms.SendKeys.SendWait("{Enter}");
                     System.Windows.Forms.SendKeys.SendWait("{Enter}");
                     System.Windows.Forms.SendKeys.SendWait("{Enter}");
                     System.Windows.Forms.SendKeys.SendWait("{Enter}");//执行现有指令（这里的多次重复主要是为了防止漏键，因为此前的调试中出现过几次漏键的情况）
                     isForwardGestureActive = false;
                     isBackGestureActive = true;
                     isLeftGestureActive = false;
                     isRightGestureActive = false;//更新四个布尔量 
                 }
             }
        }
    }
}
