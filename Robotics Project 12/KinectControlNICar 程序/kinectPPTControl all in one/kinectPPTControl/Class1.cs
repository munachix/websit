using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Kinect;

namespace kinectPPTControl
{
    class RealPosition
    {
        public double realX;
        public double realY;
        public double realZ;
        public RealPosition(Joint a)
        {
            this.realZ = a.Position.Z;
            this.realX = System.Math.Tan(Math.PI * 28.5 / 180) * a.Position.Z * a.Position.X;
            this.realY = System.Math.Tan(Math.PI * 21.5 / 180) * a.Position.Z * a.Position.Y;
        }
       
    }
}
