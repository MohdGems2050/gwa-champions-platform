import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { createClient } from "https://esm.sh/@supabase/supabase-js@2"

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  try {
    const formData = await req.formData()
    const file = formData.get('file') as File
    const competitionId = formData.get('competitionId') as string
    const studentId = formData.get('studentId') as string

    if (!file || !competitionId || !studentId) {
      return new Response(JSON.stringify({ error: 'Missing required fields' }), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
        status: 400,
      })
    }

    // Initialize Supabase with service role key from environment
    const supabase = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? ''
    )

    const timestamp = Date.now()
    const sanitizedName = file.name.replace(/[^a-z0-9.]/gi, '_').toLowerCase()
    const filePath = `${competitionId}/${studentId}/${timestamp}_${sanitizedName}`

    const { data, error } = await supabase.storage
      .from('videos')
      .upload(filePath, file, {
        contentType: file.type,
        upsert: false
      })

    if (error) throw error

    return new Response(JSON.stringify({ path: filePath }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 200,
    })
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 500,
    })
  }
})
