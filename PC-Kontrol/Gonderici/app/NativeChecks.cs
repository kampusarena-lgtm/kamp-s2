using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PcCheck {
    public static class NativeChecks {
        [StructLayout(LayoutKind.Sequential)] private struct TextureDescription {
            public uint Width, Height, MipLevels, ArraySize, Format, SampleCount, SampleQuality, Usage, BindFlags, CpuAccessFlags, MiscFlags;
        }
        [StructLayout(LayoutKind.Sequential)] private struct InitialData { public IntPtr Data; public uint RowPitch, SlicePitch; }
        [StructLayout(LayoutKind.Sequential)] private struct MappedData { public IntPtr Data; public uint RowPitch, DepthPitch; }
        [DllImport("d3d11.dll", CallingConvention=CallingConvention.StdCall)]
        private static extern int D3D11CreateDevice(IntPtr adapter, uint driverType, IntPtr software, uint flags, IntPtr levels, uint levelCount, uint sdkVersion, out IntPtr device, out uint featureLevel, out IntPtr context);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int CreateTexture(IntPtr self, ref TextureDescription desc, IntPtr initial, out IntPtr texture);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate void CopyResource(IntPtr self, IntPtr destination, IntPtr source);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate int MapResource(IntPtr self, IntPtr resource, uint subresource, uint mode, uint flags, out MappedData data);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)] private delegate void UnmapResource(IntPtr self, IntPtr resource, uint subresource);
        private static Delegate Method(IntPtr instance, int slot, Type type) {
            IntPtr vtable = Marshal.ReadIntPtr(instance);
            return Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(vtable, slot * IntPtr.Size), type);
        }
        public static byte[] Pattern() {
            byte[] data = new byte[64 * 64 * 4];
            for (int i=0; i<data.Length; i++) data[i]=(byte)((i*37+17)%251);
            return data;
        }
        public static bool Matches(byte[] expected, byte[] actual) {
            if(expected.Length!=actual.Length) return false;
            for(int i=0;i<expected.Length;i++) if(expected[i]!=actual[i]) return false;
            return true;
        }
        private static void Check(int result) { if(result<0) Marshal.ThrowExceptionForHR(result); }
        public static string Gpu() {
            if(Environment.OSVersion.Platform != PlatformID.Win32NT) return "{\"status\":\"unknown\",\"detail\":\"Direct3D requires Windows\"}";
            IntPtr device=IntPtr.Zero, context=IntPtr.Zero, source=IntPtr.Zero, staging=IntPtr.Zero, initial=IntPtr.Zero;
            GCHandle pin = new GCHandle(); bool mapped=false; uint level=0;
            try {
                // HARDWARE only: never silently fall back to Microsoft's WARP renderer.
                Check(D3D11CreateDevice(IntPtr.Zero,1,IntPtr.Zero,0,IntPtr.Zero,0,7,out device,out level,out context));
                byte[] pattern=Pattern(); pin=GCHandle.Alloc(pattern,GCHandleType.Pinned);
                InitialData data=new InitialData {Data=pin.AddrOfPinnedObject(),RowPitch=256,SlicePitch=16384};
                initial=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(InitialData))); Marshal.StructureToPtr(data,initial,false);
                TextureDescription desc=new TextureDescription {Width=64,Height=64,MipLevels=1,ArraySize=1,Format=28,SampleCount=1,Usage=0,BindFlags=8};
                CreateTexture create=(CreateTexture)Method(device,5,typeof(CreateTexture));
                Check(create(device,ref desc,initial,out source));
                desc.Usage=3;desc.BindFlags=0;desc.CpuAccessFlags=0x20000;
                Check(create(device,ref desc,IntPtr.Zero,out staging));
                ((CopyResource)Method(context,47,typeof(CopyResource)))(context,staging,source);
                MappedData memory;
                Check(((MapResource)Method(context,14,typeof(MapResource)))(context,staging,0,1,0,out memory)); mapped=true;
                byte[] actual=new byte[pattern.Length];
                for(int row=0;row<64;row++) Marshal.Copy(IntPtr.Add(memory.Data,checked((int)(row*memory.RowPitch))),actual,row*256,256);
                ((UnmapResource)Method(context,15,typeof(UnmapResource)))(context,staging,0); mapped=false;
                if(!Matches(pattern,actual)) return "{\"status\":\"fail\",\"detail\":\"Direct3D texture readback mismatch\"}";
                return "{\"status\":\"pass\",\"bytesVerified\":16384,\"featureLevel\":"+level+",\"scope\":\"Default hardware D3D11 adapter, 64x64 texture upload/copy/readback; no full GPU or VRAM stress\"}";
            } catch(Exception ex) {
                return "{\"status\":\"unknown\",\"detail\":"+Json.Quote(ex.Message)+",\"hresult\":"+ex.HResult+",\"scope\":\"Initialization/readback failure is not proof of a physical GPU fault\"}";
            } finally {
                if(mapped && context!=IntPtr.Zero && staging!=IntPtr.Zero) try { ((UnmapResource)Method(context,15,typeof(UnmapResource)))(context,staging,0); } catch { }
                if(staging!=IntPtr.Zero) Marshal.Release(staging);
                if(source!=IntPtr.Zero) Marshal.Release(source);
                if(context!=IntPtr.Zero) Marshal.Release(context);
                if(device!=IntPtr.Zero) Marshal.Release(device);
                if(initial!=IntPtr.Zero) Marshal.FreeHGlobal(initial);
                if(pin.IsAllocated) pin.Free();
            }
        }

        [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDeviceEnumerator {
            [PreserveSig] int EnumAudioEndpoints(int flow, uint mask, out IDeviceCollection devices);
            [PreserveSig] int GetDefaultAudioEndpoint(int flow, int role, out IDevice device);
        }
        [ComImport, Guid("0BD7A1BE-7A1A-44DB-8397-C0A6A4A12E61"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDeviceCollection {
            [PreserveSig] int GetCount(out uint count);
            [PreserveSig] int Item(uint index, out IDevice device);
        }
        [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDevice {
            [PreserveSig] int Activate(ref Guid id, uint context, IntPtr parameters, out IntPtr result);
            [PreserveSig] int OpenPropertyStore(uint access, out IntPtr store);
            [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
            [PreserveSig] int GetState(out uint state);
        }
        public static string AudioEndpoints() {
            if(Environment.OSVersion.Platform!=PlatformID.Win32NT) return "{\"status\":\"unknown\",\"detail\":\"CoreAudio requires Windows\"}";
            object instance=null; IDeviceCollection outputs=null; IDevice output=null;
            try {
                instance=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")));
                IDeviceEnumerator enumerator=(IDeviceEnumerator)instance;
                Check(enumerator.EnumAudioEndpoints(0,1,out outputs));uint count;Check(outputs.GetCount(out count));
                int result=enumerator.GetDefaultAudioEndpoint(0,0,out output);string id=null;
                if(result>=0&&output!=null) Check(output.GetId(out id));
                return "{\"status\":\"complete\",\"activeOutputCount\":"+count+",\"defaultOutputAvailable\":"+(result>=0?"true":"false")+",\"defaultEndpointId\":"+Json.Quote(id)+",\"scope\":\"Windows active audio endpoint only; actual headset acoustics cannot be verified unattended\"}";
            } catch(Exception ex) { return "{\"status\":\"unknown\",\"detail\":"+Json.Quote(ex.Message)+"}"; }
            finally { if(output!=null)Marshal.ReleaseComObject(output);if(outputs!=null)Marshal.ReleaseComObject(outputs);if(instance!=null)Marshal.ReleaseComObject(instance); }
        }
    }
}
